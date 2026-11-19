"""增值税 Check 检查与分录区定位。"""
import re
from decimal import Decimal, InvalidOperation

from config.config_loader import ConfigError
from log.loggingConfig import logger


TAX_RATES = ("13%", "9%", "6%")
JOURNAL_TITLES = ("账号", "借项金额", "贷项金额", "备注")


def text(value):
    return str(value).strip() if value is not None else ""


def _journal_title(value):
    """分录标题规范化：忽略空白并统一“帐/账”。"""
    return re.sub(r"\s+", "", text(value)).replace("帐", "账").replace("帳", "账")


def _rate_label(value):
    """把 0.13、13%、0.09 等统一为 13% / 9% / 6%。"""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        number = float(value)
        if 0 < number < 1:
            return f"{round(number * 100):g}%"
        if 1 <= number <= 100:
            return f"{round(number):g}%"
        return ""
    raw = text(value).replace("％", "%")
    match = re.fullmatch(r"(\d+(?:\.\d+)?)%?", raw)
    if not match:
        return ""
    number = float(match.group(1))
    if number <= 1:
        number *= 100
    return f"{round(number):g}%"


def _find_tax_columns(sheet):
    """动态定位 13%/9%/6% 对应的“税额”列。

    返回 [(税率标签, 列号)]，按从 13% 到 6% 的顺序。
    """
    header_row = None
    tax_columns = []
    for row in sheet.iter_rows():
        columns = [cell.column for cell in row if text(cell.value) == "税额"]
        if len(columns) >= 3:
            header_row = row[0].row
            tax_columns = columns
            break
    if header_row is None:
        raise ConfigError(f"{sheet.title} 未找到 13%/9%/6% 的“税额”表头")

    rate_row = header_row - 1
    rates = {}
    if rate_row >= 1:
        for cell in sheet[rate_row]:
            label = _rate_label(cell.value)
            if label in TAX_RATES and label not in rates:
                rates[label] = cell.column
    if len(rates) == 3:
        ordered = sorted(rates.items(), key=lambda item: item[1])
        mapped = []
        for index, (label, rate_column) in enumerate(ordered):
            next_rate = ordered[index + 1][1] if index + 1 < len(ordered) else None
            candidates = [column for column in tax_columns
                          if column > rate_column and (next_rate is None or column < next_rate)]
            if not candidates:
                raise ConfigError(
                    f"{sheet.title} 无法定位 {label} 对应的税额列"
                )
            mapped.append((label, candidates[0]))
        return mapped

    # 税率表头无法识别时按出现顺序回退。
    if len(tax_columns) != 3:
        raise ConfigError(
            f"{sheet.title} 找到 {len(tax_columns)} 个税额列，无法与 13%/9%/6% 对应"
        )
    return list(zip(TAX_RATES, tax_columns))


def _decimal(cell):
    value = cell.value
    if value is None or isinstance(value, bool):
        raise ConfigError(
            f"核对无法取数：{cell.parent.title}!{cell.coordinate} 为空，"
            "请确认公式已重算"
        )
    try:
        number = Decimal(str(value))
        if not number.is_finite():
            raise InvalidOperation
        return number
    except (InvalidOperation, ValueError):
        raise ConfigError(
            f"核对无法取数：{cell.parent.title}!{cell.coordinate}={value!r}，"
            "请确认公式已重算"
        )


def check_tax_differences(sheet, tolerance="5"):
    """检查“计税收入”每个 Subtotal: 下方行的 13%/9%/6% 税额差额。

    差额取 Subtotal 下一行的税额列（模板中为对计算底表的差异公式），
    绝对值不得超过容差（默认 5 元）。
    """
    try:
        limit = Decimal(str(tolerance))
        if not limit.is_finite() or limit < 0:
            raise InvalidOperation
    except (InvalidOperation, ValueError):
        raise ConfigError("核对容差必须为非负金额")

    columns = _find_tax_columns(sheet)
    subtotal_rows = []
    for row in sheet.iter_rows():
        for cell in row:
            if text(cell.value).lower() == "subtotal:":
                subtotal_rows.append(cell.row)
                break
    if not subtotal_rows:
        raise ConfigError(f"{sheet.title} 未找到“Subtotal:”行")

    unique_rows = sorted(set(subtotal_rows))
    subtotal_row = unique_rows[0]
    if len(unique_rows) > 1:
        logger.info(
            f"{sheet.title} 检测到 {len(unique_rows)} 个 Subtotal，"
            f"只校验第一个（第 {subtotal_row} 行）"
        )
    check_row = subtotal_row + 1
    if check_row > sheet.max_row:
        raise ConfigError(
            f"{sheet.title} 的 Subtotal 行 {subtotal_row} 下方没有核对行"
        )

    results = []
    for label, column in columns:
        cell = sheet.cell(check_row, column)
        difference = _decimal(cell)
        results.append({
            "核对项": f"Subtotal 第 {subtotal_row} 行 {label} 税额差额",
            "值": str(difference),
            "容差": str(limit),
            "通过": abs(difference) <= limit,
            "位置": f"{sheet.title}!{cell.coordinate}",
        })
    return results


def journal_regions(sheet):
    """定位所有“账号/借项金额/贷项金额/备注”分录区。

    标题行可不在首行、四列可不相邻，同一行可左右并排两组表头。
    返回按最左列排序的 [(表头行, {标题: 列号}, 最后账号行)]。
    数据区以最后一个有账号的行为界，中间空白账号行不跳过。
    """
    expected = {_journal_title(title) for title in JOURNAL_TITLES}
    regions = []
    for row in sheet.iter_rows():
        groups = []
        current = {}
        for cell in row:
            title = _journal_title(cell.value)
            if title not in expected:
                continue
            if title in current:
                if current:
                    groups.append(current)
                current = {}
            current[title] = cell.column
        if current:
            groups.append(current)
        for group in groups:
            if len(group) == len(expected):
                regions.append((row[0].row, group))

    if not regions:
        raise ConfigError("未找到 账号/借项金额/贷项金额/备注 分录区表头")

    result = []
    for header_row, columns in regions:
        account_column = columns["账号"]
        last_row = None
        for row in range(header_row + 1, sheet.max_row + 1):
            value = sheet.cell(row, account_column).value
            if value is not None and str(value).strip():
                last_row = row
        result.append((header_row, columns, last_row))
    result.sort(key=lambda item: min(item[1].values()))
    return result


def read_journal_rows(sheet):
    """读取所有分录区，按从左到右、左侧表格在上的顺序拼接。"""
    regions = journal_regions(sheet)
    rows = []
    for header_row, columns, last_row in regions:
        if last_row is None:
            continue
        for row in range(header_row + 1, last_row + 1):
            # 以最后账号行为界；中间空白账号行、四列全空行均保留，按空值正常录入。
            rows.append((
                sheet.cell(row, columns["账号"]).value,
                sheet.cell(row, columns["借项金额"]).value,
                sheet.cell(row, columns["贷项金额"]).value,
                sheet.cell(row, columns["备注"]).value,
            ))
    if not rows:
        raise ConfigError(f"{sheet.title} 分录区没有可录入的数据")
    description = next(
        (text(row[3]) for row in rows if text(row[3])), ""
    )
    return rows, description
