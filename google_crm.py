import re
import time
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


CRM_SCOPES = ("https://www.googleapis.com/auth/spreadsheets",)
PRODUCT_COLUMNS = {
    "iron": (10, "Железо"),
    "dojo": (11, "DOJO"),
    "oggo": (12, "OGGO"),
    "elfliq": (13, "Elfliq"),
    "cosmo": (14, "Космо"),
}
CRM_STATUSES = (
    "Холодный",
    "Ознакомились",
    "Не ознакомились",
    "отправлено КП",
    "Работаем",
    "Ждем заказ",
    "перестал заказывать",
    "Отказ/игнор/недозвон",
)


def normalize_crm_text(value: str) -> str:
    value = value.casefold().replace("ё", "е")
    return " ".join(re.sub(r"[^a-zа-я0-9@]+", " ", value).split())


@dataclass(frozen=True)
class CRMClient:
    row: int
    name: str
    region: str
    outlets: str
    status: str
    client_type: str
    contact_name: str
    phone: str
    telegram: str
    order_date: str
    call_date: str
    products: dict[str, str]
    product_notes: dict[str, str]
    note: str

    @property
    def identity(self) -> str:
        return f"{self.name}\x1f{self.telegram}\x1f{self.phone}"


class GoogleCRM:
    def __init__(self, credentials_path: Path, spreadsheet_id: str, sheet_name: str):
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        credentials = Credentials.from_service_account_file(str(credentials_path), scopes=CRM_SCOPES)
        self.service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
        self.sheet_id = self._sheet_id()
        self._client_cache: list[CRMClient] | None = None
        self._client_cache_at = 0.0

    def _sheet_id(self) -> int:
        result = self.service.spreadsheets().get(
            spreadsheetId=self.spreadsheet_id,
            fields="sheets.properties",
        ).execute()
        for sheet in result.get("sheets", []):
            properties = sheet["properties"]
            if properties["title"] == self.sheet_name:
                return int(properties["sheetId"])
        raise ValueError(f"В Google Таблице не найден лист «{self.sheet_name}».")

    def _range(self, cells: str) -> str:
        escaped = self.sheet_name.replace("'", "''")
        return f"'{escaped}'!{cells}"

    def clients(self, force: bool = False) -> list[CRMClient]:
        if not force and self._client_cache is not None and time.monotonic() - self._client_cache_at < 30:
            return self._client_cache
        values = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=self._range("A2:P"),
            valueRenderOption="FORMATTED_VALUE",
        ).execute().get("values", [])
        formats = self.service.spreadsheets().get(
            spreadsheetId=self.spreadsheet_id,
            ranges=[self._range("K2:O")],
            includeGridData=True,
            fields="sheets.data.rowData.values.effectiveFormat.backgroundColorStyle,sheets.data.rowData.values.formattedValue",
        ).execute()
        format_rows = []
        sheets = formats.get("sheets", [])
        if sheets:
            data = sheets[0].get("data", [])
            if data:
                format_rows = data[0].get("rowData", [])

        result = []
        for offset, raw in enumerate(values):
            row = list(raw) + [""] * (16 - len(raw))
            if not str(row[0]).strip():
                continue
            product_states = {}
            product_notes = {}
            format_values = format_rows[offset].get("values", []) if offset < len(format_rows) else []
            for key, (column, _) in PRODUCT_COLUMNS.items():
                product_notes[key] = str(row[column]).strip()
                cell_format = format_values[column - 10] if column - 10 < len(format_values) else {}
                color = (
                    cell_format.get("effectiveFormat", {})
                    .get("backgroundColorStyle", {})
                    .get("rgbColor", {})
                )
                red = float(color.get("red", 0))
                green = float(color.get("green", 0))
                blue = float(color.get("blue", 0))
                if green > red + 0.03 and green > blue + 0.03:
                    product_states[key] = "buy"
                elif red > green + 0.15 and red > blue + 0.15:
                    product_states[key] = "no"
                else:
                    product_states[key] = "unknown"
            result.append(CRMClient(
                row=offset + 2,
                name=str(row[0]).strip(), region=str(row[1]).strip(), outlets=str(row[2]).strip(),
                status=str(row[3]).strip(), client_type=str(row[4]).strip(),
                contact_name=str(row[5]).strip(), phone=str(row[6]).strip(),
                telegram=str(row[7]).strip(), order_date=str(row[8]).strip(),
                call_date=str(row[9]).strip(), products=product_states,
                product_notes=product_notes, note=str(row[15]).strip(),
            ))
        self._client_cache = result
        self._client_cache_at = time.monotonic()
        return result

    def invalidate(self) -> None:
        self._client_cache = None
        self._client_cache_at = 0.0

    def search(self, query: str, limit: int = 12) -> list[CRMClient]:
        query_norm = normalize_crm_text(query)
        tokens = query_norm.split()
        if not tokens:
            return []
        ranked = []
        for client in self.clients():
            main = normalize_crm_text(
                f"{client.name} {client.contact_name} {client.telegram} {client.phone} {client.region}"
            )
            extended = normalize_crm_text(f"{main} {client.note}")
            words = extended.split()
            token_scores = []
            for token in tokens:
                if token in words:
                    token_scores.append(1.0)
                elif token in extended:
                    token_scores.append(0.94)
                else:
                    token_scores.append(max(
                        (SequenceMatcher(None, token, word).ratio() for word in words), default=0
                    ))
            if all(score >= (0.78 if len(token) <= 3 else 0.62) for token, score in zip(tokens, token_scores)):
                score = sum(token_scores) / len(token_scores)
                if query_norm in main:
                    score += 0.2
                ranked.append((score, client))
        ranked.sort(key=lambda item: (item[0], -item[1].row), reverse=True)
        return [client for _, client in ranked[:limit]]

    def client(self, row: int, expected_identity: str = "") -> CRMClient | None:
        clients = self.clients()
        direct = next((client for client in clients if client.row == row), None)
        if direct and (not expected_identity or direct.identity == expected_identity):
            return direct
        if expected_identity:
            return next((client for client in clients if client.identity == expected_identity), None)
        return direct

    def _update_value(self, cell: str, value: str) -> None:
        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=self._range(cell),
            valueInputOption="USER_ENTERED",
            body={"values": [[value]]},
        ).execute()

    def append_note(self, row: int, expected_identity: str, text: str, when: datetime) -> CRMClient:
        client = self.client(row, expected_identity)
        if not client:
            raise ValueError("Клиент не найден: возможно, таблицу изменили или отсортировали.")
        entry = f"{when.strftime('%d.%m.%Y %H:%M')}\n{text.strip()}"
        value = f"{client.note}\n\n{entry}" if client.note else entry
        self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": [
                    {"range": self._range(f"P{client.row}"), "values": [[value]]},
                    {"range": self._range(f"J{client.row}"), "values": [[when.strftime('%d.%m.%Y')]]},
                ],
            },
        ).execute()
        self.invalidate()
        return client

    def update_status(self, row: int, expected_identity: str, status: str) -> CRMClient:
        if status not in CRM_STATUSES:
            raise ValueError("Неизвестный статус клиента.")
        client = self.client(row, expected_identity)
        if not client:
            raise ValueError("Клиент не найден: возможно, таблицу изменили или отсортировали.")
        self._update_value(f"D{client.row}", status)
        self.invalidate()
        return client

    def mark_order_today(self, row: int, expected_identity: str, when: datetime) -> CRMClient:
        client = self.client(row, expected_identity)
        if not client:
            raise ValueError("Клиент не найден: возможно, таблицу изменили или отсортировали.")
        self._update_value(f"I{client.row}", when.strftime("%d.%m.%Y"))
        self.invalidate()
        return client

    def set_product(
        self, row: int, expected_identity: str, product: str, state: str, reason: str = ""
    ) -> CRMClient:
        if product not in PRODUCT_COLUMNS or state not in {"buy", "no", "unknown"}:
            raise ValueError("Некорректная товарная группа или отметка.")
        if state == "no" and not reason.strip():
            raise ValueError("Для красной отметки нужно указать причину.")
        client = self.client(row, expected_identity)
        if not client:
            raise ValueError("Клиент не найден: возможно, таблицу изменили или отсортировали.")
        column, _ = PRODUCT_COLUMNS[product]
        text = client.product_notes.get(product, "")
        if state == "no":
            text = reason.strip()
        elif state == "buy" and client.products.get(product) == "no":
            text = ""
        elif state == "unknown":
            text = ""
        colors = {
            "buy": {"red": 0.0, "green": 1.0, "blue": 0.0},
            "no": {"red": 1.0, "green": 0.0, "blue": 0.0},
        }
        cell = {"userEnteredValue": {"stringValue": text}, "userEnteredFormat": {}}
        if state in colors:
            cell["userEnteredFormat"]["backgroundColor"] = colors[state]
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": [{"updateCells": {
                "range": {
                    "sheetId": self.sheet_id,
                    "startRowIndex": client.row - 1,
                    "endRowIndex": client.row,
                    "startColumnIndex": column,
                    "endColumnIndex": column + 1,
                },
                "rows": [{"values": [cell]}],
                "fields": "userEnteredValue,userEnteredFormat.backgroundColor",
            }}]},
        ).execute()
        self.invalidate()
        return client
