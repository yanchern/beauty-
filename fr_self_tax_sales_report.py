from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


EU_COUNTRIES = {
    "AT",
    "BE",
    "BG",
    "HR",
    "CY",
    "CZ",
    "DE",
    "DK",
    "EE",
    "ES",
    "FI",
    "FR",
    "GR",
    "HU",
    "IE",
    "IT",
    "LT",
    "LU",
    "LV",
    "MT",
    "NL",
    "PL",
    "PT",
    "RO",
    "SE",
    "SI",
    "SK",
}


@dataclass(frozen=True)
class RuleContext:
    country_code: str


@dataclass(frozen=True)
class RuleSpec:
    rule_id: str
    logic_group: str
    logic_bucket: str
    description: str
    matcher: Callable[[dict[str, str], RuleContext], bool]


@dataclass(frozen=True)
class CountryConfig:
    code: str
    slug: str
    name_zh: str
    title: str
    description: str
    sales_report_label: str
    logic_pdf_label: str
    excluded_transaction_types: tuple[str, ...]
    rules: tuple[RuleSpec, ...]


@dataclass(frozen=True)
class MatchResult:
    rule_id: str
    logic_group: str
    logic_bucket: str
    description: str


@dataclass
class MatchedRow:
    source_row_number: int
    rule_id: str
    logic_group: str
    logic_bucket: str
    rule_description: str
    transaction_event_id: str
    activity_period: str
    marketplace: str
    transaction_type: str
    tax_collection_responsibility: str
    tax_reporting_scheme: str
    seller_sku: str
    asin: str
    item_description: str
    qty: str
    total_activity_value_amt_vat_incl: Decimal
    transaction_currency_code: str
    price_of_items_vat_rate_percent: str
    buyer_vat_number: str
    sale_depart_country: str
    sale_arrival_country: str
    taxable_jurisdiction: str
    vat_calculation_imputation_country: str
    invoice_url: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="根据国家税金逻辑，从 Amazon 销售报告中提取自行缴税订单并输出 Excel 汇总。"
    )
    parser.add_argument("csv_path", type=Path, help="季度销售报告 CSV 路径")
    parser.add_argument(
        "--country",
        choices=("fr", "es"),
        default="fr",
        help="国家代码。fr=法国，es=西班牙。",
    )
    parser.add_argument(
        "--logic-pdf-path",
        type=Path,
        default=None,
        help="税金逻辑 PDF 路径，仅用于在汇总页记录来源。",
    )
    parser.add_argument(
        "--output-path",
        type=Path,
        default=None,
        help="输出 Excel 路径；默认在 CSV 同目录生成 *_自行缴税销售额汇总.xlsx。",
    )
    return parser.parse_args()


def decimal_or_zero(raw: str) -> Decimal:
    value = (raw or "").strip()
    if not value:
        return Decimal("0")
    try:
        return Decimal(value.replace(",", ""))
    except InvalidOperation as exc:
        raise ValueError(f"无法解析金额: {raw!r}") from exc


def normalized(row: dict[str, str], key: str) -> str:
    return (row.get(key) or "").strip()


def upper_value(row: dict[str, str], key: str) -> str:
    return normalized(row, key).upper()


def normalized_transaction_type(row: dict[str, str]) -> str:
    raw = upper_value(row, "TRANSACTION_TYPE")
    return re.sub(r"[_\-\s]+", " ", raw).strip()


def is_blank(raw: str) -> bool:
    return raw.strip() == ""


def is_zero(raw: str) -> bool:
    return decimal_or_zero(raw) == 0


def is_blank_or_zero(raw: str) -> bool:
    value = raw.strip()
    return value == "" or is_zero(value)


def is_non_zero(raw: str) -> bool:
    value = raw.strip()
    return value != "" and not is_zero(value)


def is_seller_responsible(row: dict[str, str]) -> bool:
    return upper_value(row, "TAX_COLLECTION_RESPONSIBILITY") == "SELLER"


def base_allowed(row: dict[str, str], ctx: RuleContext) -> bool:
    country = COUNTRY_CONFIGS[ctx.country_code]
    return normalized_transaction_type(row) not in country.excluded_transaction_types


def fr_rule_part_2(row: dict[str, str], ctx: RuleContext) -> bool:
    return (
        base_allowed(row, ctx)
        and is_seller_responsible(row)
        and is_non_zero(normalized(row, "PRICE_OF_ITEMS_VAT_RATE_PERCENT"))
        and not is_blank(normalized(row, "BUYER_VAT_NUMBER"))
        and upper_value(row, "SALE_DEPART_COUNTRY") == "FR"
        and upper_value(row, "SALE_ARRIVAL_COUNTRY") in EU_COUNTRIES
    )


def fr_rule_part_3_strict(row: dict[str, str], ctx: RuleContext) -> bool:
    return (
        base_allowed(row, ctx)
        and is_seller_responsible(row)
        and is_non_zero(normalized(row, "PRICE_OF_ITEMS_VAT_RATE_PERCENT"))
        and is_blank(normalized(row, "BUYER_VAT_NUMBER"))
        and upper_value(row, "SALE_DEPART_COUNTRY") in EU_COUNTRIES
        and upper_value(row, "SALE_ARRIVAL_COUNTRY") == "FR"
    )


def fr_rule_part_3_missing(row: dict[str, str], ctx: RuleContext) -> bool:
    return (
        base_allowed(row, ctx)
        and is_seller_responsible(row)
        and is_blank_or_zero(normalized(row, "PRICE_OF_ITEMS_VAT_RATE_PERCENT"))
        and is_blank(normalized(row, "BUYER_VAT_NUMBER"))
        and upper_value(row, "SALE_DEPART_COUNTRY") in EU_COUNTRIES
        and upper_value(row, "SALE_ARRIVAL_COUNTRY") == "FR"
    )


def es_rule_part_1(row: dict[str, str], ctx: RuleContext) -> bool:
    return (
        base_allowed(row, ctx)
        and is_seller_responsible(row)
        and is_non_zero(normalized(row, "PRICE_OF_ITEMS_VAT_RATE_PERCENT"))
        and not is_blank(normalized(row, "BUYER_VAT_NUMBER"))
        and upper_value(row, "SALE_DEPART_COUNTRY") == "ES"
        and upper_value(row, "SALE_ARRIVAL_COUNTRY") in EU_COUNTRIES
    )


def es_rule_part_2(row: dict[str, str], ctx: RuleContext) -> bool:
    return (
        base_allowed(row, ctx)
        and is_seller_responsible(row)
        and is_non_zero(normalized(row, "PRICE_OF_ITEMS_VAT_RATE_PERCENT"))
        and is_blank(normalized(row, "BUYER_VAT_NUMBER"))
        and upper_value(row, "SALE_DEPART_COUNTRY") in EU_COUNTRIES
        and upper_value(row, "SALE_ARRIVAL_COUNTRY") == "ES"
    )


def es_rule_part_3(row: dict[str, str], ctx: RuleContext) -> bool:
    return (
        base_allowed(row, ctx)
        and is_seller_responsible(row)
        and is_blank(normalized(row, "PRICE_OF_ITEMS_VAT_RATE_PERCENT"))
        and is_blank(normalized(row, "BUYER_VAT_NUMBER"))
        and upper_value(row, "SALE_DEPART_COUNTRY") in EU_COUNTRIES
        and upper_value(row, "SALE_ARRIVAL_COUNTRY") == "ES"
    )


COUNTRY_CONFIGS: dict[str, CountryConfig] = {
    "fr": CountryConfig(
        code="fr",
        slug="france",
        name_zh="法国",
        title="法国季度申报税金计算",
        description=(
            "按法国税金逻辑提取自行缴税订单，固定包含第 2 部分、"
            "第 3 部分严格口径，以及你确认过的第 3 部分遗漏订单口径。"
        ),
        sales_report_label="法国销售报告 CSV",
        logic_pdf_label="法国税金逻辑 PDF",
        excluded_transaction_types=("COMMINGLING BUY", "COMMINGLING SELLER"),
        rules=(
            RuleSpec(
                rule_id="FR_SELF_TAX_P2",
                logic_group="自行缴税第2部分",
                logic_bucket="规则2-1_FR出发_买家税号有效",
                description=(
                    "F 不属于 COMMINGLING_BUY / COMMINGLING_SELLER；CQ=SELLER；"
                    "AE 非空且非 0；CA 有值；BP=FR；BQ 属于欧盟国家。"
                ),
                matcher=fr_rule_part_2,
            ),
            RuleSpec(
                rule_id="FR_SELF_TAX_P3_STRICT",
                logic_group="自行缴税第3部分",
                logic_bucket="规则2-2_目的国FR_买家税号空白_AE非零",
                description=(
                    "F 不属于 COMMINGLING_BUY / COMMINGLING_SELLER；CQ=SELLER；"
                    "AE 非空且非 0；CA 空白；BP 属于欧盟国家；BQ=FR。"
                ),
                matcher=fr_rule_part_3_strict,
            ),
            RuleSpec(
                rule_id="FR_SELF_TAX_P3_MISSING",
                logic_group="自行缴税第3部分",
                logic_bucket="规则2-2_目的国FR_买家税号空白_AE空白或零_按遗漏订单纳入",
                description=(
                    "固定补入遗漏订单口径：F 不属于 COMMINGLING_BUY / COMMINGLING_SELLER；"
                    "CQ=SELLER；AE 为空白或为 0；CA 空白；BP 属于欧盟国家；BQ=FR。"
                ),
                matcher=fr_rule_part_3_missing,
            ),
        ),
    ),
    "es": CountryConfig(
        code="es",
        slug="spain",
        name_zh="西班牙",
        title="西班牙季度申报税金计算",
        description=(
            "按西班牙计税原则提取自行缴税订单，覆盖第一部分、第二部分、第三部分三段规则。"
        ),
        sales_report_label="西班牙销售报告 CSV",
        logic_pdf_label="西班牙税金逻辑 PDF",
        excluded_transaction_types=("COMMINGLING BUY",),
        rules=(
            RuleSpec(
                rule_id="ES_SELF_TAX_P1",
                logic_group="自行缴税第一部分",
                logic_bucket="规则2-1_ES出发_买家税号有效",
                description=(
                    "F 不等于 COMMINGLING BUY；CQ=SELLER；AE 非空且非 0；"
                    "CA 有值；BP=ES；BQ 属于欧盟国家。"
                ),
                matcher=es_rule_part_1,
            ),
            RuleSpec(
                rule_id="ES_SELF_TAX_P2",
                logic_group="自行缴税第二部分",
                logic_bucket="规则2-2_目的国ES_买家税号空白_AE非零",
                description=(
                    "F 不等于 COMMINGLING BUY；CQ=SELLER；AE 非空且非 0；"
                    "CA 空白；BP 属于欧盟国家；BQ=ES。"
                ),
                matcher=es_rule_part_2,
            ),
            RuleSpec(
                rule_id="ES_SELF_TAX_P3",
                logic_group="自行缴税第三部分",
                logic_bucket="规则2-3_目的国ES_买家税号空白_AE空白",
                description=(
                    "F 不等于 COMMINGLING BUY；CQ=SELLER；AE 为空白；"
                    "CA 空白；BP 属于欧盟国家；BQ=ES。"
                ),
                matcher=es_rule_part_3,
            ),
        ),
    ),
}


def get_country_config(country_code: str) -> CountryConfig:
    try:
        return COUNTRY_CONFIGS[country_code]
    except KeyError as exc:
        raise ValueError(f"不支持的国家代码: {country_code}") from exc


def get_group_order(country: CountryConfig) -> list[str]:
    ordered: list[str] = []
    for rule in country.rules:
        if rule.logic_group not in ordered:
            ordered.append(rule.logic_group)
    return ordered


def match_self_tax_logic(
    row: dict[str, str],
    ctx: RuleContext,
    country: CountryConfig,
) -> MatchResult | None:
    for rule in country.rules:
        if rule.matcher(row, ctx):
            return MatchResult(
                rule_id=rule.rule_id,
                logic_group=rule.logic_group,
                logic_bucket=rule.logic_bucket,
                description=rule.description,
            )
    return None


def iter_matched_rows(csv_path: Path, country: CountryConfig) -> list[MatchedRow]:
    matched_rows: list[MatchedRow] = []
    ctx = RuleContext(country_code=country.code)
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for source_row_number, row in enumerate(reader, start=2):
            match = match_self_tax_logic(row, ctx, country)
            if not match:
                continue

            matched_rows.append(
                MatchedRow(
                    source_row_number=source_row_number,
                    rule_id=match.rule_id,
                    logic_group=match.logic_group,
                    logic_bucket=match.logic_bucket,
                    rule_description=match.description,
                    transaction_event_id=row.get("TRANSACTION_EVENT_ID", ""),
                    activity_period=row.get("ACTIVITY_PERIOD", ""),
                    marketplace=row.get("MARKETPLACE", ""),
                    transaction_type=row.get("TRANSACTION_TYPE", ""),
                    tax_collection_responsibility=row.get("TAX_COLLECTION_RESPONSIBILITY", ""),
                    tax_reporting_scheme=row.get("TAX_REPORTING_SCHEME", ""),
                    seller_sku=row.get("SELLER_SKU", ""),
                    asin=row.get("ASIN", ""),
                    item_description=row.get("ITEM_DESCRIPTION", ""),
                    qty=row.get("QTY", ""),
                    total_activity_value_amt_vat_incl=decimal_or_zero(
                        row.get("TOTAL_ACTIVITY_VALUE_AMT_VAT_INCL", "")
                    ),
                    transaction_currency_code=row.get("TRANSACTION_CURRENCY_CODE", ""),
                    price_of_items_vat_rate_percent=row.get("PRICE_OF_ITEMS_VAT_RATE_PERCENT", ""),
                    buyer_vat_number=row.get("BUYER_VAT_NUMBER", ""),
                    sale_depart_country=row.get("SALE_DEPART_COUNTRY", ""),
                    sale_arrival_country=row.get("SALE_ARRIVAL_COUNTRY", ""),
                    taxable_jurisdiction=row.get("TAXABLE_JURISDICTION", ""),
                    vat_calculation_imputation_country=row.get(
                        "VAT_CALCULATION_IMPUTATION_COUNTRY", ""
                    ),
                    invoice_url=row.get("INVOICE_URL", ""),
                )
            )
    return matched_rows


def autosize_worksheet(ws) -> None:
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)


def summarize_by_group(matched_rows: list[MatchedRow]) -> dict[str, Decimal]:
    totals: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    for row in matched_rows:
        totals[row.logic_group] += row.total_activity_value_amt_vat_incl
    return dict(totals)


def summarize_by_rule(
    matched_rows: list[MatchedRow],
    country: CountryConfig,
) -> dict[str, dict[str, object]]:
    summary: dict[str, dict[str, object]] = {}
    for rule in country.rules:
        summary[rule.rule_id] = {
            "logic_group": rule.logic_group,
            "logic_bucket": rule.logic_bucket,
            "description": rule.description,
            "count": 0,
            "amount": Decimal("0"),
        }

    for row in matched_rows:
        bucket = summary[row.rule_id]
        bucket["count"] = int(bucket["count"]) + 1
        bucket["amount"] = Decimal(bucket["amount"]) + row.total_activity_value_amt_vat_incl
    return summary


def build_summary_sheet(
    wb: Workbook,
    matched_rows: list[MatchedRow],
    csv_path: Path,
    logic_pdf_path: Path | None,
    country: CountryConfig,
) -> None:
    ws = wb.active
    ws.title = "汇总"

    currency_codes = sorted(
        {row.transaction_currency_code for row in matched_rows if row.transaction_currency_code}
    )
    totals_by_group = summarize_by_group(matched_rows)
    total_sales = sum(
        (row.total_activity_value_amt_vat_incl for row in matched_rows),
        start=Decimal("0"),
    )
    sale_only_total = sum(
        (
            row.total_activity_value_amt_vat_incl
            for row in matched_rows
            if row.transaction_type.strip().upper() == "SALE"
        ),
        start=Decimal("0"),
    )

    summary_rows: list[tuple[str, object]] = [
        ("国家", country.name_zh),
        ("输入销售报告", str(csv_path)),
        ("税金逻辑来源", str(logic_pdf_path) if logic_pdf_path else "未提供"),
        ("命中自行缴税记录数", len(matched_rows)),
    ]
    for group in get_group_order(country):
        summary_rows.append((f"{group}销售额(含税)", float(totals_by_group.get(group, Decimal("0")))))
    summary_rows.extend(
        [
            ("自行缴税合计销售额(含税)", float(total_sales)),
            ("仅SALE交易销售额总额(含税)", float(sale_only_total)),
            ("币种", ", ".join(currency_codes) if currency_codes else ""),
        ]
    )

    ws.append(("指标", "值"))
    for row in summary_rows:
        ws.append(row)

    ws["A1"].font = Font(bold=True)
    ws["B1"].font = Font(bold=True)
    for row_index in range(2, ws.max_row + 1):
        if isinstance(ws.cell(row=row_index, column=2).value, (int, float)):
            ws.cell(row=row_index, column=2).number_format = "0.00"

    autosize_worksheet(ws)
    ws.freeze_panes = "A2"


def build_rule_sheet(
    wb: Workbook,
    matched_rows: list[MatchedRow],
    country: CountryConfig,
) -> None:
    ws = wb.create_sheet("规则说明")
    ws.append(
        [
            "rule_id",
            "logic_group",
            "logic_bucket",
            "matched_rows",
            "matched_amount",
            "description",
        ]
    )
    for cell in ws[1]:
        cell.font = Font(bold=True)

    rule_summary = summarize_by_rule(matched_rows, country)
    for rule in country.rules:
        ws.append(
            [
                rule.rule_id,
                rule.logic_group,
                rule.logic_bucket,
                rule_summary[rule.rule_id]["count"],
                float(Decimal(rule_summary[rule.rule_id]["amount"])),
                rule.description,
            ]
        )

    for row_index in range(2, ws.max_row + 1):
        ws.cell(row=row_index, column=5).number_format = "0.00"

    autosize_worksheet(ws)
    ws.freeze_panes = "A2"


def build_detail_sheet(wb: Workbook, matched_rows: list[MatchedRow]) -> None:
    ws = wb.create_sheet("命中明细")
    headers = [
        "source_row_number",
        "rule_id",
        "logic_group",
        "logic_bucket",
        "rule_description",
        "transaction_event_id",
        "activity_period",
        "marketplace",
        "transaction_type",
        "tax_collection_responsibility",
        "tax_reporting_scheme",
        "seller_sku",
        "asin",
        "item_description",
        "qty",
        "total_activity_value_amt_vat_incl",
        "transaction_currency_code",
        "price_of_items_vat_rate_percent",
        "buyer_vat_number",
        "sale_depart_country",
        "sale_arrival_country",
        "taxable_jurisdiction",
        "vat_calculation_imputation_country",
        "invoice_url",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in matched_rows:
        ws.append(
            [
                row.source_row_number,
                row.rule_id,
                row.logic_group,
                row.logic_bucket,
                row.rule_description,
                row.transaction_event_id,
                row.activity_period,
                row.marketplace,
                row.transaction_type,
                row.tax_collection_responsibility,
                row.tax_reporting_scheme,
                row.seller_sku,
                row.asin,
                row.item_description,
                row.qty,
                float(row.total_activity_value_amt_vat_incl),
                row.transaction_currency_code,
                row.price_of_items_vat_rate_percent,
                row.buyer_vat_number,
                row.sale_depart_country,
                row.sale_arrival_country,
                row.taxable_jurisdiction,
                row.vat_calculation_imputation_country,
                row.invoice_url,
            ]
        )

    amount_column = headers.index("total_activity_value_amt_vat_incl") + 1
    for row_index in range(2, ws.max_row + 1):
        ws.cell(row=row_index, column=amount_column).number_format = "0.00"

    autosize_worksheet(ws)
    ws.freeze_panes = "A2"


def build_workbook(
    matched_rows: list[MatchedRow],
    csv_path: Path,
    logic_pdf_path: Path | None,
    country: CountryConfig,
) -> Workbook:
    wb = Workbook()
    build_summary_sheet(wb, matched_rows, csv_path, logic_pdf_path, country)
    build_rule_sheet(wb, matched_rows, country)
    build_detail_sheet(wb, matched_rows)
    return wb


def generate_self_tax_report(
    csv_path: Path,
    country_code: str,
    logic_pdf_path: Path | None = None,
    output_path: Path | None = None,
) -> dict[str, object]:
    country = get_country_config(country_code)
    csv_path = csv_path.expanduser().resolve()
    logic_pdf_path = logic_pdf_path.expanduser().resolve() if logic_pdf_path else None

    if not csv_path.exists():
        raise FileNotFoundError(f"未找到销售报告: {csv_path}")

    if output_path is None:
        output_path = csv_path.with_name(f"{csv_path.stem}_自行缴税销售额汇总.xlsx")
    output_path = output_path.expanduser().resolve()

    matched_rows = iter_matched_rows(csv_path, country)
    workbook = build_workbook(matched_rows, csv_path, logic_pdf_path, country)
    workbook.save(output_path)

    totals_by_group = summarize_by_group(matched_rows)
    total_sales = sum(
        (row.total_activity_value_amt_vat_incl for row in matched_rows),
        start=Decimal("0"),
    )

    return {
        "country_code": country.code,
        "country_name": country.name_zh,
        "country_slug": country.slug,
        "output_path": output_path,
        "matched_rows": matched_rows,
        "row_count": len(matched_rows),
        "group_order": get_group_order(country),
        "group_totals": {key: totals_by_group.get(key, Decimal("0")) for key in get_group_order(country)},
        "total_sales": total_sales,
    }


def main() -> None:
    args = parse_args()
    result = generate_self_tax_report(
        csv_path=args.csv_path,
        country_code=args.country,
        logic_pdf_path=args.logic_pdf_path,
        output_path=args.output_path,
    )

    print(f"国家: {result['country_name']}")
    print(f"输出文件: {result['output_path']}")
    print(f"命中自行缴税记录数: {result['row_count']}")
    for group in result["group_order"]:
        print(f"{group}销售额(含税): {result['group_totals'][group]:.2f}")
    print(f"自行缴税订单销售额总额(含税): {result['total_sales']:.2f}")


if __name__ == "__main__":
    main()
