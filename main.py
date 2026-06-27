import argparse
import csv
import os
import re
import sys
import yaml
from dotenv import load_dotenv

load_dotenv()

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import pickle

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_PATH = os.path.join(os.path.dirname(__file__), "token.pickle")
CREDENTIALS_PATH = os.path.join(os.path.dirname(__file__), "credentials.json")
DEFAULT_FILTERS_PATH = os.path.join(os.path.dirname(__file__), "filters.yaml")

OUTPUT_HEADERS = [
    "Статус регистрации",
    "Сайт актуален?",
    "ID регистрации",
    "ФИО",
    "Email",
    "Telegram",
    "Волна",
    "Билет",
    "Статус",
    "Тип оплаты",
    "Полная сумма",
    "Скидка",
    "Итого",
    "Время регистрации",
    "Время подтверждения",
    "Получатель (имя)",
    "Получатель (банк)",
    "Реквизиты перевода",
]


def load_filters(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def parse_csv(path: str) -> list[dict]:
    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        return list(reader)


def normalize_tg(handle: str) -> str:
    return handle.lstrip("@").lower().strip()


def deduplicate(rows: list[dict]) -> list[dict]:
    """Keep first entry per ФИО (CSV is ordered newest→oldest, so first = most recent)."""
    seen: dict[str, dict] = {}
    for row in rows:
        name = row.get("ФИО", "").strip()
        if name and name not in seen:
            seen[name] = row
    return list(seen.values())


def is_blacklisted(handle: str, filters: dict) -> bool:
    tg = normalize_tg(handle)
    bl = filters.get("blacklist", {})
    if tg in {h.lower() for h in bl.get("exact", [])}:
        return True
    if any(tg.startswith(p.lower()) for p in bl.get("prefixes", [])):
        return True
    if any(re.fullmatch(p, tg) for p in bl.get("patterns", [])):
        return True
    return False


def filter_blacklist(rows: list[dict], filters: dict) -> list[dict]:
    return [row for row in rows if not is_blacklisted(row.get("Telegram", ""), filters)]


def _to_int(val: str):
    try:
        return int(val.strip())
    except (ValueError, AttributeError):
        return val


def build_output_rows(rows: list[dict], filters: dict) -> list[tuple]:
    status_map = filters.get("status_map", {})
    excluded = set(filters.get("status_filter", {}).get("exclude", []))
    out = []
    for row in rows:
        status_raw = row.get("Статус", "").strip()
        if status_raw in excluded:
            continue
        label = status_map.get(status_raw, status_raw)
        record = {
            "Статус регистрации": label,
            "Сайт актуален?": True,
            "ID регистрации": row.get("ID регистрации", ""),
            "ФИО": row.get("ФИО", ""),
            "Email": row.get("Email", ""),
            "Telegram": row.get("Telegram", ""),
            "Волна": row.get("Волна", ""),
            "Билет": row.get("Билет", ""),
            "Статус": status_raw,
            "Тип оплаты": row.get("Тип оплаты", ""),
            "Полная сумма": _to_int(row.get("Полная сумма", "")),
            "Скидка": _to_int(row.get("Скидка", "")),
            "Итого": _to_int(row.get("Итого", "")),
            "Время регистрации": row.get("Время регистрации", ""),
            "Время подтверждения": row.get("Время подтверждения", ""),
            "Получатель (имя)": row.get("Получатель (имя)", ""),
            "Получатель (банк)": row.get("Получатель (банк)", ""),
            "Реквизиты перевода": row.get("Реквизиты перевода", ""),
        }
        out.append((record, status_raw))
    return out


def write_xlsx(output_rows: list[tuple], path: str, filters: dict):
    wb = Workbook()
    ws = wb.active
    ws.title = "Регистрации"

    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    header_font = Font(bold=True)

    for col, header in enumerate(OUTPUT_HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font

    status_map = filters.get("status_map", {})
    xlsx_colors = filters.get("xlsx_colors", {})

    for row_idx, (record, status_raw) in enumerate(output_rows, start=2):
        label = status_map.get(status_raw, status_raw)
        color_hex = xlsx_colors.get(label)
        for col, header in enumerate(OUTPUT_HEADERS, start=1):
            val = record.get(header, "")
            cell = ws.cell(row=row_idx, column=col, value=val if val is not True else True)
            if header == "Статус регистрации" and color_hex:
                cell.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

    wb.save(path)


def get_google_creds() -> Credentials:
    creds = None
    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                print(
                    f"Файл credentials.json не найден по пути {CREDENTIALS_PATH}.\n"
                    "Скачай его из Google Cloud Console (OAuth 2.0 Desktop) и положи рядом со скриптом.",
                    file=sys.stderr,
                )
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "wb") as f:
            pickle.dump(creds, f)
    return creds


def push_to_sheets(output_rows: list[tuple]):
    creds = get_google_creds()
    service = build("sheets", "v4", credentials=creds)
    sheet = service.spreadsheets()

    sheet.values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
    ).execute()

    values = [OUTPUT_HEADERS]
    for record, _ in output_rows:
        row = []
        for h in OUTPUT_HEADERS:
            v = record.get(h, "")
            row.append("TRUE" if v is True else ("FALSE" if v is False else v))
        values.append(row)

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    requests = []

    meta = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == SHEET_NAME:
            sheet_id = s["properties"]["sheetId"]
            break

    if sheet_id is None:
        print(f"Лист '{SHEET_NAME}' не найден в таблице.", file=sys.stderr)
        return

    requests.append({
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 1,
                "endRowIndex": len(output_rows) + 1,
                "startColumnIndex": 1,
                "endColumnIndex": 2,
            },
            "rule": {
                "condition": {"type": "BOOLEAN"},
                "showCustomUi": True,
            },
        }
    })

    if requests:
        sheet.batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()

    print(f"Данные ({len(output_rows)} строк) загружены в Google Sheets: '{SHEET_NAME}'")


def main():
    parser = argparse.ArgumentParser(description="Обработка регистраций из CSV")
    parser.add_argument("input", help="Путь к входному CSV файлу")
    parser.add_argument("output", help="Путь к выходному XLSX файлу")
    parser.add_argument(
        "--filters",
        default=DEFAULT_FILTERS_PATH,
        help="Путь к YAML файлу с фильтрами (по умолчанию: filters.yaml)",
    )
    parser.add_argument(
        "--no-sheets",
        action="store_true",
        help="Не загружать данные в Google Sheets",
    )
    args = parser.parse_args()

    filters = load_filters(args.filters)
    sort_order = filters.get("sort_order", {})

    rows = parse_csv(args.input)
    rows = filter_blacklist(rows, filters)
    rows = deduplicate(rows)
    output_rows = build_output_rows(rows, filters)
    output_rows.sort(key=lambda x: sort_order.get(x[1], 99))

    write_xlsx(output_rows, args.output, filters)
    print(f"XLSX сохранён: {args.output} ({len(output_rows)} строк)")

    if not args.no_sheets:
        push_to_sheets(output_rows)


if __name__ == "__main__":
    main()
