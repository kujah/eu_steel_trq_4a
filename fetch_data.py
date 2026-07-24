from __future__ import annotations

import json
import re
import ssl
from dataclasses import dataclass
from datetime import datetime, timezone
from html import unescape
from pathlib import Path
from typing import Any
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Font


ORDER_CODES = [
    "099832",
    "099833",
    "099834",
    "099835",
    "099836",
    "099505",
    "099605",
    "099714",
    "099718",
    "099716",
    "099715",
    "099717",
]

EXPORTER_MAP = {
    "099832": "Viet Nam",
    "099833": "Taiwan",
    "099834": "Turkiye",
    "099835": "India",
    "099836": "Korea",
    "099505": "FTA Quota - CSQ",
    "099605": "Other countries",
    "099714": "FTA Quota - Other countries",
    "099718": "United Kingdom",
    "099716": "Japan",
    "099715": "Egypt",
    "099717": "South Africa",
}

QUOTA_SECTION = "4.A"
PRODUCT_GROUP = "Metallic Coated Sheets"
EXPORTER_ORDER = [
    "Other countries",
    "Korea",
    "Viet Nam",
    "Taiwan",
    "Turkiye",
    "India",
    "FTA Quota - CSQ",
    "United Kingdom",
    "Japan",
    "Egypt",
    "South Africa",
    "FTA Quota - Other countries",
]

BASE_URL = "https://ec.europa.eu/taxation_customs/dds2/taric"
LIST_URL = (
    BASE_URL
    + "/quota_list.jsp?Lang=en&Code={code}&Year=2026&Expand=true&Offset=0"
)
DETAIL_URL = BASE_URL + "/quota_tariff_details.jsp?Lang=en&StartDate={start_date}&Code={code}"
DATA_DIR = Path(__file__).resolve().parent / "public" / "data"
OUT_PATH = DATA_DIR / "orders.json"
XLSX_PATH = DATA_DIR / "eu-steel-trq-4a-dashboard.xlsx"


@dataclass
class Quantity:
    value: float | None
    unit: str | None
    raw: str


def fetch_text(url: str) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    ssl_context = ssl._create_unverified_context()
    with urlopen(request, timeout=30, context=ssl_context) as response:
        raw = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="replace")


def squash_whitespace(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_html_text(fragment: str) -> str:
    fragment = re.sub(r"<br\s*/?>", "\n", fragment, flags=re.IGNORECASE)
    fragment = re.sub(r"</div\s*>", "\n", fragment, flags=re.IGNORECASE)
    fragment = re.sub(r"<div[^>]*>", "", fragment, flags=re.IGNORECASE)
    fragment = re.sub(r"</?a[^>]*>", "", fragment, flags=re.IGNORECASE)
    fragment = re.sub(r"<[^>]+>", " ", fragment)
    lines = [squash_whitespace(unescape(line)) for line in fragment.splitlines()]
    return "\n".join(line for line in lines if line)


def parse_quantity(text: str) -> Quantity:
    raw = squash_whitespace(text)
    match = re.match(r"^([0-9]+(?:[.,][0-9]+)?)\s+(.+)$", raw)
    if not match:
        number_match = re.match(r"^([0-9]+(?:[.,][0-9]+)?)$", raw)
        if number_match:
            return Quantity(float(number_match.group(1).replace(",", "")), None, raw)
        return Quantity(None, None, raw)
    value = float(match.group(1).replace(",", ""))
    unit = match.group(2).strip()
    return Quantity(value, unit, raw)


def find_first(pattern: str, text: str) -> str | None:
    match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    return match.group(1) if match else None


def parse_table_rows(container_html: str) -> dict[str, str]:
    rows: dict[str, str] = {}
    row_pattern = re.compile(
        r"<tr[^>]*class=\"ecl-table__row\"[^>]*>(.*?)</tr>",
        flags=re.IGNORECASE | re.DOTALL,
    )
    cell_pattern = re.compile(r"<td[^>]*>(.*?)</td>", flags=re.IGNORECASE | re.DOTALL)
    for row_html in row_pattern.findall(container_html):
        cells = cell_pattern.findall(row_html)
        if len(cells) < 2:
            continue
        label_html, value_html = cells[0], cells[1]
        label = clean_html_text(label_html)
        label = re.sub(r"\(\s*indicative\s*\)", "", label, flags=re.IGNORECASE)
        label = squash_whitespace(label)
        value = clean_html_text(value_html)
        rows[label] = value
    return rows


def parse_list_page(code: str, html: str) -> dict[str, Any]:
    tbody_match = re.search(
        r"<tbody[^>]*class=\"ecl-table__body\"[^>]*>(.*?)</tbody>",
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not tbody_match:
        raise RuntimeError(f"No list body found for order number {code}")
    row_match = re.search(
        r"<tr[^>]*class=\"ecl-table__row\"[^>]*>(.*?)</tr>",
        tbody_match.group(1),
        flags=re.IGNORECASE | re.DOTALL,
    )
    if not row_match:
        raise RuntimeError(f"No list row found for order number {code}")
    row_html = row_match.group(1)

    cells = re.findall(r"<td[^>]*>(.*?)</td>", row_html, flags=re.IGNORECASE | re.DOTALL)
    if len(cells) < 6:
        raise RuntimeError(f"Unexpected list row structure for order number {code}")

    detail_href = find_first(r'href="([^"]*quota_tariff_details\.jsp[^"]+)"', row_html)
    start_date = find_first(r"StartDate=([0-9]{4}-[0-9]{2}-[0-9]{2})", detail_href or "")
    last_update = find_first(r"Last TARIC update:&nbsp;([0-9]{2}-[0-9]{2}-[0-9]{4})", html)

    return {
        "order_number": clean_html_text(cells[0]),
        "origins_summary": clean_html_text(cells[1]),
        "start_date": clean_html_text(cells[2]),
        "end_date": clean_html_text(cells[3]),
        "balance": parse_quantity(clean_html_text(cells[4])).__dict__,
        "detail_url": detail_href,
        "detail_start_date": start_date,
        "source_last_taric_update": last_update,
    }


def parse_detail_page(code: str, html: str) -> dict[str, Any]:
    start = html.find('<div id="quotaDetailsMarkedUpContainer"')
    marker = html.find("Associated TARIC code", start)
    end = html.find("</table>", marker)
    container = html[start:end] if start >= 0 and marker >= 0 and end >= 0 else None
    if not container:
        raise RuntimeError(f"No detail container found for order number {code}")

    rows = parse_table_rows(container)
    last_update = find_first(r"Last TARIC update:&nbsp;([0-9]{2}-[0-9]{2}-[0-9]{4})", html)

    initial_amount = parse_quantity(rows.get("Initial amount", ""))
    current_amount = parse_quantity(rows.get("Amount", ""))
    balance = parse_quantity(rows.get("Balance", ""))

    used_value = None
    if initial_amount.value is not None and balance.value is not None:
        used_value = round(initial_amount.value - balance.value, 2)

    utilization_pct = None
    if initial_amount.value and used_value is not None:
        utilization_pct = round((used_value / initial_amount.value) * 100, 2)

    return {
        "order_number": rows.get("Order number", code),
        "validity_period": rows.get("Validity period", ""),
        "origin": rows.get("Origin", ""),
        "initial_amount": initial_amount.__dict__,
        "current_amount": current_amount.__dict__,
        "balance": balance.__dict__,
        "used_amount": {
            "value": used_value,
            "unit": initial_amount.unit or balance.unit,
            "raw": "" if used_value is None else f"{used_value}",
        },
        "utilization_pct": utilization_pct,
        "exhaustion_date": rows.get("Exhaustion date", ""),
        "critical": rows.get("Critical", ""),
        "last_import_date": rows.get("Last import date", ""),
        "last_allocation_date": rows.get("Last allocation date", ""),
        "awaiting_allocation": rows.get("Total awaiting allocation", ""),
        "awaiting_allocation_mt": to_mt(
            parse_quantity(rows.get("Total awaiting allocation", "")).__dict__
        ),
        "blocking_period": rows.get("Blocking period", ""),
        "suspension_period": rows.get("Suspension period", ""),
        "allocated_pct_last_allocation": rows.get("Allocated percentage at the last allocation", ""),
        "associated_taric_codes": rows.get("Associated TARIC code", "").splitlines(),
        "source_last_taric_update": last_update,
    }


def to_mt(quantity: dict[str, Any]) -> dict[str, Any]:
    value = quantity.get("value")
    if value is None:
        return {"value": None, "unit": "MT", "raw": ""}
    converted = round(value / 1000, 2)
    return {"value": converted, "unit": "MT", "raw": f"{converted:.2f} MT"}


def build_record(code: str) -> dict[str, Any]:
    list_html = fetch_text(LIST_URL.format(code=code))
    list_data = parse_list_page(code, list_html)
    if not list_data["detail_start_date"]:
        raise RuntimeError(f"Missing detail start date for order number {code}")

    detail_html = fetch_text(
        DETAIL_URL.format(code=code, start_date=list_data["detail_start_date"])
    )
    detail_data = parse_detail_page(code, detail_html)

    source_last_update = detail_data["source_last_taric_update"] or list_data["source_last_taric_update"]
    return {
        "order_number": code,
        "quota_section": QUOTA_SECTION,
        "product_group": PRODUCT_GROUP,
        "exporter": EXPORTER_MAP.get(code, list_data["origins_summary"]),
        "origins_summary": list_data["origins_summary"],
        "start_date": list_data["start_date"],
        "end_date": list_data["end_date"],
        "detail_url": list_data["detail_url"],
        "source_last_taric_update": source_last_update,
        "initial_amount_mt": to_mt(detail_data["initial_amount"]),
        "used_amount_mt": to_mt(detail_data["used_amount"]),
        "balance_mt": to_mt(detail_data["balance"]),
        **detail_data,
    }


def build_excel(payload: dict[str, Any]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TRQ 4A Dashboard"

    headers = [
        "Quota Section",
        "Product Group",
        "Order Number",
        "Exporter",
        "Start Date",
        "End Date",
        "Initial (MT)",
        "Used (MT)",
        "Balance (MT)",
        "Utilization (%)",
        "Critical",
        "Last Allocation Date",
        "Awaiting Allocation (MT)",
        "Detail URL",
    ]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for item in payload["items"]:
        sheet.append(
            [
                item["quota_section"],
                item["product_group"],
                item["order_number"],
                item["exporter"],
                item["start_date"],
                item["end_date"],
                item["initial_amount_mt"]["value"],
                item["used_amount_mt"]["value"],
                item["balance_mt"]["value"],
                item["utilization_pct"],
                item["critical"],
                item["last_allocation_date"],
                item["awaiting_allocation_mt"]["value"],
                item["detail_url"],
            ]
        )

    for row in sheet.iter_rows(min_row=2, min_col=7, max_col=10):
        for cell in row:
            cell.number_format = "0.00"

    widths = {
        "A": 14,
        "B": 24,
        "C": 14,
        "D": 24,
        "E": 14,
        "F": 14,
        "G": 14,
        "H": 14,
        "I": 14,
        "J": 14,
        "K": 12,
        "L": 20,
        "M": 18,
        "N": 80,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    workbook.save(XLSX_PATH)


def main() -> None:
    records = [build_record(code) for code in ORDER_CODES]
    exporter_rank = {name: index for index, name in enumerate(EXPORTER_ORDER)}
    records.sort(key=lambda item: exporter_rank.get(item["exporter"], 999))
    report_period = ""
    if records:
        report_period = f'{records[0]["start_date"]} - {records[0]["end_date"]}'

    payload = {
        "title": "EU STEEL TRQ 4A 소진 현황",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "source": "https://ec.europa.eu/taxation_customs/dds2/taric/quota_consultation.jsp?Lang=en",
        "year_filter": 2026,
        "order_numbers": ORDER_CODES,
        "quota_section": QUOTA_SECTION,
        "product_group": PRODUCT_GROUP,
        "report_period": report_period,
        "items": records,
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    build_excel(payload)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
