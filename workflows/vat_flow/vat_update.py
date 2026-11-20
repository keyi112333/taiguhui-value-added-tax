"""增值税底稿生成：导入明细账、写入上月留抵报表、更新会计期间。"""
import re
import shutil
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

from config.config_loader import ConfigError
from log.loggingConfig import logger


EXCEL_BASE_DATE = datetime(1899, 12, 30)
REQUIRED_LEDGER_HEADERS = ("帐号", "单据 公司", "分类帐类型 1 金额")


def excel_serial(value):
    """把 Python 日期转换为 Excel 日期序列号，避免 COM 时区偏差。"""
    if isinstance(value, datetime):
        delta = value - EXCEL_BASE_DATE
        return delta.days + delta.seconds / 86400
    if isinstance(value, date):
        return (value - EXCEL_BASE_DATE.date()).days
    return value


def _excel_value(value):
    """文本以 = 开头时加单引号，避免 JDE 数据被 Excel 当成公式。"""
    value = excel_serial(value)
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def normalize_header(value):
    """标题规范化：忽略空白、下划线、连字符与括号，并统一“帐/账”。"""
    text = "" if value is None else str(value).strip().lower()
    text = text.replace("帐", "账").replace("帳", "账")
    return re.sub(r"[\s_\-+/()（）]+", "", text)


def normalize_company(value):
    """把公司代码统一为 5 位字符串（02501）。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text.zfill(5) if text.isdigit() else text.upper()


def _open_book(app, filename, read_only):
    """新启的 Excel 实例偶发“Open 方法无效”，重试可覆盖启动竞态。"""
    last = None
    for _ in range(3):
        try:
            return app.Workbooks.Open(
                str(filename), UpdateLinks=0, ReadOnly=read_only,
                IgnoreReadOnlyRecommended=True, AddToMru=False,
            )
        except Exception as exc:
            last = exc
            time.sleep(1)
    raise ConfigError(f"Excel 无法打开 {filename}：{last}")


def read_ledger(path, sheet_name=None, header_row=1):
    """读取 JDE 导出的明细账，返回 (标题列表, 数据行列表)。"""
    path = Path(path)
    if not path.is_file():
        raise ConfigError(f"明细账文件不存在：{path}")
    # JDE 导出的 xlsx 可能缺少完整的 dimension 元数据；read_only 模式会
    # 按错误范围只读到第一列，因此这里使用普通模式读取。
    book = openpyxl.load_workbook(path, data_only=True)
    try:
        if sheet_name and sheet_name in book.sheetnames:
            sheet = book[sheet_name]
        else:
            sheet = book[book.sheetnames[0]]
        headers = []
        for values in sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True):
            headers = [value for value in values]
            break
        rows = []
        for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            if all(value is None or (isinstance(value, str) and not value.strip())
                   for value in values):
                continue
            rows.append(list(values))
        return headers, rows
    finally:
        book.close()


def filter_ledger(rows, headers, company_code):
    """筛选“单据 公司”等于公司代码的数据行。"""
    index = {normalize_header(name): position for position, name in enumerate(headers)
             if normalize_header(name)}
    key = normalize_header("单据 公司")
    if key not in index:
        raise ConfigError("明细账缺少“单据 公司”列，无法按公司筛选")
    position = index[key]
    target = normalize_company(company_code)
    selected = []
    for row in rows:
        value = row[position] if position < len(row) else None
        if normalize_company(value) == target:
            selected.append(row)
    if not selected:
        raise ConfigError(
            f"明细账中没有“单据 公司”为 {target} 的数据；"
            f"请检查导出条件或公司代码配置"
        )
    return selected


def _sheet_names(book):
    return [book.Worksheets(index).Name for index in range(1, book.Sheets.Count + 1)]


def _require_sheet(book, name, description):
    names = _sheet_names(book)
    if name not in names:
        raise ConfigError(f"{description}缺少工作表：{name}；当前工作表：{names}")
    return book.Worksheets(name)


def _resolve_sheet(book, names, description):
    existing = _sheet_names(book)
    for name in names:
        if name in existing:
            return book.Worksheets(name)
    raise ConfigError(f"{description}缺少工作表：{'/'.join(names)}；当前工作表：{existing}")


def _target_headers(sheet, header_row, last_col):
    return [sheet.Cells(header_row, column).Value for column in range(1, last_col + 1)]


def _clear_ledger_filter(sheet):
    """去掉 JDE明细账 上的筛选，保留标题行并展开被隐藏的行。

    模板若带着筛选保存，写入只会作用于可见区域，清除与写入都会错位，
    新数据也可能被旧筛选条件隐藏，因此写入前先取消筛选。
    """
    if sheet.FilterMode:
        sheet.ShowAllData()
        logger.info(f"JDE明细账存在筛选结果，已全部显示：{sheet.Name}")
    if sheet.AutoFilterMode:
        sheet.AutoFilterMode = False
        logger.info(f"JDE明细账已取消自动筛选（保留标题行）：{sheet.Name}")
    # 取消筛选不会自动展开之前被筛选隐藏的行。
    sheet.UsedRange.EntireRow.Hidden = False


def _write_ledger(sheet, header_row, source_headers, rows, company_code):
    """按标题映射把明细账写入模板的 JDE明细账 页。"""
    _clear_ledger_filter(sheet)
    used = sheet.UsedRange
    last_col = max(used.Column + used.Columns.Count - 1, len(source_headers))
    target_headers = _target_headers(sheet, header_row, last_col)
    source_index = {normalize_header(name): position
                    for position, name in enumerate(source_headers)
                    if normalize_header(name)}
    target_map = {}
    missing = []
    for column, title in enumerate(target_headers, 1):
        key = normalize_header(title)
        if not key:
            continue
        if key in source_index:
            target_map[column] = source_index[key]
        else:
            missing.append(str(title))
    if missing:
        raise ConfigError(f"模板 JDE明细账 的标题在导出明细账中不存在：{missing}")

    target_keys = {normalize_header(title) for title in target_headers}
    lacking = [name for name in REQUIRED_LEDGER_HEADERS
               if normalize_header(name) not in target_keys]
    if lacking:
        raise ConfigError(f"模板 JDE明细账 缺少必需标题：{lacking}")
    extra = [str(title) for title in source_headers
             if normalize_header(title) and normalize_header(title) not in
             {normalize_header(value) for value in target_headers}]
    if extra:
        logger.warning(f"明细账存在模板未使用的列，已忽略：{extra}")

    selected = filter_ledger(rows, source_headers, company_code)
    logger.info(
        f"明细账筛选完成：原始 {len(rows)} 行，单据公司 {normalize_company(company_code)} "
        f"保留 {len(selected)} 行"
    )
    if len(selected) + header_row > 58534:
        logger.warning(
            "明细账行数超过模板部分公式的固定范围（58534 行），"
            "请检查增值税计算底表/计税收入公式是否需要扩展行号"
        )

    old_last_row = used.Row + used.Rows.Count - 1
    if old_last_row > header_row:
        sheet.Range(
            sheet.Cells(header_row + 1, 1),
            sheet.Cells(old_last_row, last_col),
        ).ClearContents()

    if not selected:
        return 0

    aligned = []
    for row in selected:
        values = []
        for column in range(1, last_col + 1):
            position = target_map.get(column)
            value = row[position] if position is not None and position < len(row) else None
            values.append(_excel_value(value))
        aligned.append(values)

    end_row = header_row + len(aligned)
    # 整列都是文本时先设为文本格式，避免 "02501"、单据号等被 Excel 转成数字。
    original_formats = {
        column: sheet.Cells(header_row + 1, column).NumberFormat
        for column in range(1, last_col + 1)
    }
    text_columns = []
    for column in range(1, last_col + 1):
        values = [row[column - 1] for row in aligned]
        if values and all(value is None or isinstance(value, str) for value in values):
            text_columns.append(column)
    for column in text_columns:
        sheet.Range(
            sheet.Cells(header_row + 1, column),
            sheet.Cells(end_row, column),
        ).NumberFormat = "@"

    # 分批写入，避免一次性向 COM 传递过大的二维数组。
    chunk_size = 2000
    for offset in range(0, len(aligned), chunk_size):
        chunk = aligned[offset:offset + chunk_size]
        start_row = header_row + 1 + offset
        chunk_end = start_row + len(chunk) - 1
        sheet.Range(
            sheet.Cells(start_row, 1),
            sheet.Cells(chunk_end, last_col),
        ).Value2 = tuple(tuple(row) for row in chunk)

    # 写入后恢复模板原有格式；文本列的文本类型仍然保留前导零。
    for column in range(1, last_col + 1):
        sheet.Range(
            sheet.Cells(header_row + 1, column),
            sheet.Cells(end_row, column),
        ).NumberFormat = original_formats[column]
    logger.info(f"JDE明细账已写入：{len(aligned)} 行 × {last_col} 列")
    return len(aligned)


def _find_report_header(sheet, max_rows=60):
    """定位 R55GL1005 标题行：A 列为 Cat23 的那一行。"""
    for row in range(1, max_rows + 1):
        value = sheet.Cells(row, 1).Value
        if isinstance(value, str) and value.strip().lower() == "cat23":
            return row
    raise ConfigError(f"{sheet.Name} 未找到 R55GL1005 标题行（Cat23）")


def _last_report_row(sheet, header_row, last_row):
    """从表尾向上跳过空行，返回最后一行的数据/汇总行号。"""
    row = last_row
    while row > header_row:
        values = [sheet.Cells(row, column).Value for column in (1, 2, 5)]
        if any(value is not None and str(value).strip() for value in values):
            return row
        row -= 1
    return header_row


def _fill_ending_balance(sheet, header_row, last_row):
    """在 L 列填充 I 列减 J 列（期末借贷差额），汇总行除外。"""
    data_last = _last_report_row(sheet, header_row, last_row)
    label = sheet.Cells(data_last, 5).Value
    if isinstance(label, str) and label.strip().lower().startswith("total"):
        data_last -= 1
    first = header_row + 1
    if data_last < first:
        return 0
    formulas = tuple((f"=I{row}-J{row}",) for row in range(first, data_last + 1))
    sheet.Range(
        sheet.Cells(first, 12),
        sheet.Cells(data_last, 12),
    ).Formula = formulas
    return len(formulas)


def _copy_report(app, book, report_path, target_name):
    """把 R55GL1005 报表整表写入“上月留抵”页，并在 L 列补 I-J 差额。"""
    report_path = Path(report_path)
    if not report_path.is_file():
        raise ConfigError(f"R55GL1005 报表不存在：{report_path}")
    target = _resolve_sheet(book, [target_name, "上月留底", "上月留抵"], "工作簿")
    report_book = _open_book(app, report_path, True)
    try:
        source = report_book.Worksheets(1)
        used = source.UsedRange
        try:
            target.UsedRange.UnMerge()
        except Exception:
            pass
        target.UsedRange.ClearContents()
        used.Copy(target.Range(used.Address))
        app.CutCopyMode = False
        logger.info(
            f"R55GL1005 已写入 {target.Name}：{used.Address}"
            f"（{used.Rows.Count} 行 × {used.Columns.Count} 列）"
        )
        # 报表只有 A-J 列，复制前的 ClearContents 会清掉 L 列公式，需重新填充。
        header_row = _find_report_header(target)
        rows = _fill_ending_balance(
            target, header_row, used.Row + used.Rows.Count - 1
        )
        logger.info(f"L 列期末差额（I-J）已写入：{rows} 行")
    finally:
        report_book.Close(SaveChanges=False)


def _find_label(sheet, label, max_rows=60, max_columns=40):
    for row in range(1, max_rows + 1):
        for column in range(1, max_columns + 1):
            value = sheet.Cells(row, column).Value
            if isinstance(value, str) and value.strip() == label:
                return sheet.Cells(row, column)
    return None


def _find_contains(sheet, text, max_rows=60, max_columns=40):
    for row in range(1, max_rows + 1):
        for column in range(1, max_columns + 1):
            value = sheet.Cells(row, column).Value
            if isinstance(value, str) and text in value:
                return sheet.Cells(row, column)
    return None


def _set_accounting_period(sheet, first_day):
    cell = _find_label(sheet, "会计期间：")
    if cell is None:
        raise ConfigError(f"{sheet.Name} 未找到“会计期间：”单元格")
    target = sheet.Cells(cell.Row, cell.Column + 1)
    target.Value2 = excel_serial(first_day)
    target.NumberFormat = "yyyy/m/d"
    logger.info(f"会计期间已更新为 {first_day:%Y/%m/%d}：{sheet.Name}!{target.Address}")


def _set_tax_period_text(sheet, first_day, last_day):
    cell = _find_contains(sheet, "税款所属期间：")
    if cell is None:
        raise ConfigError(f"{sheet.Name} 未找到“税款所属期间：”单元格")
    text = str(cell.Value or "")
    prefix = text.split("税款所属期间：")[0] + "税款所属期间："
    value = prefix + f"{first_day:%Y年%m月%d日}至{last_day:%Y年%m月%d日}"
    cell.Value2 = value
    logger.info(f"税款所属期间已更新：{sheet.Name}!{cell.Address} = {value.strip()}")


def build_workbook(template, destination, ledger_path, report_path, period_first,
                   company_code, ledger_sheet="JDE明细账", ledger_header_row=1,
                   calc_sheet="增值税计算底表", main_sheet="主表",
                   last_month_sheet="上月留抵", engine="Excel.Application"):
    """复制模板并完成明细账导入、上月留抵写入与会计期间更新。

    所有写入由 Excel COM 完成，保留模板中的公式、图表、数据验证与外部链接。
    保存前统一全量重算，确保缓存值与公式一致。
    """
    import pythoncom
    import win32com.client

    template = Path(template)
    destination = Path(destination)
    if not template.is_file():
        raise ConfigError(f"模板文件不存在：{template}")
    if destination.resolve() == template.resolve():
        raise ConfigError(f"输出文件不能与模板重合：{destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()
        logger.info(f"已删除同名输出文件：{destination}")

    # 先读取明细账，缩短 Excel 会话占用时间。
    source_headers, source_rows = read_ledger(ledger_path, None, ledger_header_row)

    pythoncom.CoInitialize()
    app = book = None
    try:
        app = win32com.client.DispatchEx(engine)
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3

        if template.suffix.lower() == ".xlsx":
            shutil.copy2(template, destination)
            book = _open_book(app, destination, False)
        else:
            # 旧版 .xls 模板先另存为 .xlsx，保留公式、格式与工作表。
            book = _open_book(app, template, False)
            book.SaveAs(str(destination), 51)
            logger.info(f"模板已转换为工作簿格式：{destination}")
            if not destination.is_file():
                raise ConfigError(
                    f"模板转换后未生成文件：{destination}；"
                    f"Excel 实际路径：{book.FullName}"
                )
        if book.ReadOnly:
            raise ConfigError("输出文件被占用，无法写入")

        sheet = _require_sheet(book, ledger_sheet, "模板")
        written = _write_ledger(sheet, ledger_header_row, source_headers,
                                source_rows, company_code)
        if written == 0:
            raise ConfigError("明细账没有可写入的数据")

        _copy_report(app, book, report_path, last_month_sheet)

        calc = _require_sheet(book, calc_sheet, "模板")
        _set_accounting_period(calc, period_first)
        # 下月首日减一天即本月最后一天。
        next_month = (period_first.replace(day=28) + timedelta(days=4)).replace(day=1)
        month_end = next_month - timedelta(days=1)
        main = _require_sheet(book, main_sheet, "模板")
        _set_tax_period_text(main, period_first, month_end)

        # 保存前重算，避免写入明细账/上月留抵后沿用模板旧缓存。
        app.CalculateFullRebuild()
        book.Save()
        actual = Path(str(book.FullName))
        if actual.resolve() != destination.resolve():
            logger.warning(
                f"Excel 实际保存路径与目标不一致：{actual}，"
                f"重新另存为：{destination}"
            )
            book.SaveAs(str(destination), 51)
            book.Save()
            actual = Path(str(book.FullName))
        if not destination.is_file() or destination.stat().st_size == 0:
            raise ConfigError(
                f"生成文件未落盘：{destination}；Excel 实际路径：{actual}"
            )
        logger.info(f"增值税计税表已生成：{destination}")
        return destination, written
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def write_batch_number(path, sheet_name, batch, engine="Excel.Application"):
    """把 JDE 批号写回“批号：”右侧单元格。"""
    if not batch:
        return
    import pythoncom
    import win32com.client

    path = Path(path)
    pythoncom.CoInitialize()
    app = book = None
    try:
        app = win32com.client.DispatchEx(engine)
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        book = _open_book(app, path, False)
        if book.ReadOnly:
            raise ConfigError("生成文件被占用，无法回写批号")
        sheet = _require_sheet(book, sheet_name, "工作簿")
        cell = _find_label(sheet, "批号：")
        if cell is None:
            raise ConfigError(f"{sheet_name} 未找到“批号：”单元格")
        target = sheet.Cells(cell.Row, cell.Column + 1)
        target.Value2 = str(batch)
        book.Save()
        logger.info(f"JDE 批号 {batch} 已回写：{sheet_name}!{target.Address}")
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
