"""增值税模块二至五：导出、生成底稿、Check 检查与分录录入。"""
import calendar
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl

from config.config_loader import ConfigError
from log.loggingConfig import logger
from utils.helpers import get_timestamp
from web.JDE.jde_function import (
    create_journal_entry,
    download_gl_movement,
    download_r55gl1005,
)
from workflows.vat_flow.template_check import (
    convert_legacy_excel_to_xlsx,
    recalculate_file,
)
from workflows.vat_flow.vat_check import check_tax_differences, read_journal_rows
from workflows.vat_flow.vat_update import build_workbook, write_batch_number


def configured_period(value):
    """会计期间配置值；空白（含只有空格）视为未配置，按自然月返回今天。"""
    text = "" if value is None else str(value).strip()
    return text or date.today()


def month_bounds(period):
    """会计期间自然月第一天和最后一天。"""
    if isinstance(period, (datetime, date)):
        year, month = period.year, period.month
    else:
        text = "" if period is None else str(period).strip()
        match = re.fullmatch(r"(\d{4})[-/年]?(\d{1,2})月?", text)
        if not match:
            raise ConfigError(
                f"会计期间格式不正确，应为 YYYY-MM 或 YYYY/MM 或 YYYYMM 等可识别格式；"
                f"当前值：{text}"
            )
        year, month = int(match.group(1)), int(match.group(2))
    if not (1 <= month <= 12):
        raise ConfigError(f"会计期间月份不正确，应为 1-12；当前值：{month}")
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def shift_month(first_day, offset):
    """按自然月偏移，返回目标月份的第一天；offset 可为负数。"""
    total = first_day.year * 12 + (first_day.month - 1) + offset
    year, month = divmod(total, 12)
    return date(year, month + 1, 1)


def format_jde_date(value, label):
    if isinstance(value, (datetime, date)):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, pattern).strftime("%m/%d/%Y")
        except ValueError:
            continue
    raise ConfigError(
        f"{label} 日期格式不正确，应为 YYYY-MM-DD 或 MM/DD/YYYY 等可识别格式；当前值：{text}"
    )


def flag(value):
    """把配置值转换为布尔值。"""
    if isinstance(value, str):
        return value.strip().lower() in ("true", "1", "yes", "y", "是")
    return bool(value)


class VatFlow:
    def __init__(self, context):
        self.context = context
        self.config = context.config

    def get(self, key, default=None):
        return self.config.get(f"增值税配置.{key}", default)

    def accounting_period(self):
        """会计期间首日：显式配置则用配置值，空白按自然月（当月）。"""
        first, _ = month_bounds(configured_period(self.get("会计期间")))
        return first

    def ledger_dates(self):
        """明细账日期：会计期间当月及往后三个月，共四个月。"""
        first, _ = month_bounds(configured_period(self.get("会计期间")))
        _, last = month_bounds(shift_month(first, 3))
        return format_jde_date(first, "明细账起始日期"), format_jde_date(last, "明细账结束日期")

    def export_ledger(self):
        """模块二：从 JDE 导出明细账，返回下载文件路径。"""
        required = ("明细账自公司", "明细账至公司")
        missing = [name for name in required
                   if self.get(name) is None or not str(self.get(name)).strip()]
        if missing:
            raise ConfigError(f"模块二缺少导出参数：{', '.join(missing)}")
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；请先登录 JDE")
        start_date, end_date = self.ledger_dates()

        destination = self.get("明细账文件")
        if destination:
            destination = Path(destination)
            if not destination.is_absolute():
                destination = self.config.get_path("output") / destination
        else:
            destination = self.config.get_path("output") / f"增值税明细账 {get_timestamp()}"
        path = download_gl_movement(
            page,
            ledger_type=self.get("明细账分类帐类型", "AA"),
            from_company=self.get("明细账自公司"),
            to_company=self.get("明细账至公司"),
            start_date=start_date,
            end_date=end_date,
            destination=destination,
        )
        self.context.set("jde_ledger_path", path)
        return path

    def export_r55gl1005(self):
        """导出本会计期间的 R55GL1005，写入模板的“上月留抵”页。"""
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；无法导出 R55GL1005")

        company = self.get("明细账自公司")
        if company is None or not str(company).strip():
            raise ConfigError("导出 R55GL1005 缺少 增值税配置.明细账自公司")

        destination = (
            self.context.get("jde_r55gl1005_path")
            or self.get("R55GL1005文件")
        )
        if destination:
            destination = Path(destination)
            if not destination.is_absolute():
                destination = self.config.get_path("input") / destination
        else:
            destination = self.config.get_path("output") / (
                f"R55GL1005_{get_timestamp()}.xlsx"
            )
        destination = Path(destination)
        if destination.suffix.lower() != ".xlsx":
            destination = destination.with_suffix(".xlsx")

        period = self.accounting_period()
        logger.info(f"R55GL1005 取数期间：{period:%Y-%m}")
        raw_path = download_r55gl1005(
            page,
            company=company,
            destination=destination.with_suffix(".xls"),
            period=period,
        )
        convert_legacy_excel_to_xlsx(raw_path, destination)
        logger.info(f"R55GL1005 已转换为工作簿格式：{destination}")
        self.context.set("jde_r55gl1005_path", destination)
        return destination

    def resolve_template(self):
        """解析共享盘/输入目录中的增值税计税表模板。"""
        input_dir = self.config.get_path("input")
        template_name = self.get("模板文件")
        if template_name:
            template = Path(template_name)
            if not template.is_absolute():
                template = input_dir / template
            if not template.is_file():
                raise ConfigError(f"模板文件不存在：{template}")
            return template

        candidates = []
        for pattern in ("*增值税计税表*.xlsx", "*增值税计税表*.xls"):
            candidates.extend(
                path for path in input_dir.glob(pattern)
                if not path.name.startswith("~$")
            )
        if len(candidates) != 1:
            raise ConfigError(
                "输入目录需有唯一包含“增值税计税表”的 xlsx/xls 模板，或配置 增值税配置.模板文件"
            )
        return candidates[0]

    def build_template(self, ledger_path, report_path):
        """模块三：生成增值税计税表并导入明细账、上月留抵与会计期间。"""
        template = self.resolve_template()
        company = self.get("公司")
        if company is None or not str(company).strip() or re.search(
            r'[<>:"/\\|?*\x00-\x1f]', str(company)
        ):
            raise ConfigError("增值税配置.公司 必填且不能包含文件名非法字符")

        period = f"{self.accounting_period():%Y%m}"
        output_dir = Path(self.get("输出目录") or self.config.get_path("output"))
        if not output_dir.is_absolute():
            output_dir = self.config.get_path("output") / output_dir
        destination = (
            output_dir / f"增值税计税表 {str(company).strip()} {period}.xlsx"
        ).resolve()
        if destination.resolve() in (template.resolve(), Path(ledger_path).resolve()):
            raise ConfigError(f"输出文件不能与输入重合：{destination}")

        path, rows = build_workbook(
            template=template,
            destination=destination,
            ledger_path=ledger_path,
            report_path=report_path,
            period_first=self.accounting_period(),
            company_code=str(self.get("明细账自公司") or ""),
            ledger_sheet=self.get("明细账工作表", "JDE明细账"),
            ledger_header_row=int(self.get("明细账标题行", 1)),
            calc_sheet=self.get("计算底表工作表", "增值税计算底表"),
            main_sheet=self.get("主表工作表", "主表"),
            last_month_sheet=self.get("上月留抵工作表", "上月留抵"),
            engine=str(self.get("重算引擎", "Excel.Application")).strip(),
        )
        self.context.set("vat_path", path)
        self.context.set("jde_ledger_rows", rows)
        return path

    def check(self, path):
        """模块四：重算并检查“计税收入”Subtotal 下方 13%/9%/6% 税额差额。"""
        self.context.set("vat_check_status", "running")
        try:
            recalculate_file(path, str(self.get("重算引擎", "Excel.Application")).strip())
            book = openpyxl.load_workbook(path, data_only=True)
            try:
                name = self.get("计税收入工作表", "计税收入")
                if name not in book.sheetnames:
                    raise ConfigError(f"工作簿缺少工作表：{name}")
                results = check_tax_differences(
                    book[name],
                    tolerance=self.get("核对容差", "5"),
                )
            finally:
                book.close()
            passed = all(item["通过"] for item in results)
            self.context.set("vat_check_results", results)
            self.context.set("vat_check_status", "passed" if passed else "failed")
            for item in results:
                logger.info(
                    f"[{'通过' if item['通过'] else '不通过'}] {item['核对项']}："
                    f"差额={item['值']} 容差={item['容差']} 位置={item['位置']}"
                )
            if not passed:
                failed = "、".join(item["核对项"] for item in results if not item["通过"])
                message = f"表格核对失败：{failed}"
                logger.error(message)
                self.context.set("vat_pending_reason", message)
                return False
        except Exception:
            self.context.set("vat_check_status", "failed")
            raise
        return True

    def journal_rows(self, path):
        book = openpyxl.load_workbook(path, data_only=True)
        try:
            name = self.get("分录工作表", "增值税计算底表")
            if name not in book.sheetnames:
                raise ConfigError(f"工作簿缺少工作表：{name}")
            return read_journal_rows(book[name])
        finally:
            book.close()

    def write_journal(self, path):
        """模块五：把分录区整块粘贴录入 JDE，并回写批号。"""
        browser = self.context.get_browser("JDE")
        page = browser.get_page() if browser else None
        if page is None:
            raise RuntimeError("JDE 浏览器未打开；请先登录 JDE")
        rows, description = self.journal_rows(path)
        batch = create_journal_entry(
            page,
            company=self.get("分录公司", "02501"),
            description=description,
            gl_date=format_jde_date(date.today(), "总帐日期"),
            rows=rows,
            discard=not flag(self.get("分录保存", False)),
        )
        self.context.set("jde_journal_batch", batch)
        self.context.set("jde_journal_rows", len(rows))
        write_batch_number(
            path,
            self.get("分录工作表", "增值税计算底表"),
            batch,
            engine=str(self.get("重算引擎", "Excel.Application")).strip(),
        )
        return batch

    def run(self):
        self.context.set("vat_status", "running")
        try:
            ledger_path = self.export_ledger()
            report_path = self.export_r55gl1005()
            path = self.build_template(ledger_path, report_path)
            decision = self.check(path)
            self.context.set("journal_required", decision)
            if not decision:
                # 核对不通过：仅记录并提示，不抛异常、不进入分录。
                self.context.set("vat_status", "failed")
                return
            self.write_journal(path)
            self.context.set("vat_status", "completed")
        except Exception:
            self.context.set("vat_status", "failed")
            raise
