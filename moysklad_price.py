"""Read-only MoySklad integration used to build the common price list."""

from __future__ import annotations

import json
import gzip
import ssl
import time
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Callable
from urllib.parse import urlencode, urlparse
from urllib.error import HTTPError, URLError
from urllib.request import HTTPRedirectHandler, HTTPSHandler, Request, build_opener, urlopen

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Font, PatternFill


API_ROOT = "https://api.moysklad.ru/api/remap/1.2"
DEFAULT_STORES = ("Мордор", "Годзибасы", "Жможики")
BONUS_CATEGORIES = (
    "Космо Аромы/жижи",
    "Эльфлик + LM Аромы/жижи",
    "Разки",
    "OGGO Аромы/жижи",
    "Железо",
    "Не учитывать",
)
MAX_BONUS_CATEGORY_CHOICES = 21


class MoySkladError(RuntimeError):
    def __init__(self, message: str, *, status: int | None = None) -> None:
        super().__init__(message)
        self.status = status


@dataclass(frozen=True)
class MoySkladPriceResult:
    path: Path
    item_count: int
    group_count: int
    store_names: tuple[str, ...]
    cash_price_type: str
    cashless_price_type: str


def _canonical_href(value: str) -> str:
    return value.split("?", 1)[0].rstrip("/")


def _money(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        # Monetary values in JSON API are returned in minor currency units.
        result = Decimal(str(value)) / Decimal(100)
    except Exception:
        return None
    return result if result > 0 else None


class MoySkladClient:
    """A deliberately read-only client: it exposes GET operations only."""

    def __init__(
        self,
        token: str,
        *,
        timeout: int = 90,
        opener: Callable = urlopen,
    ) -> None:
        token = token.strip()
        if not token:
            raise MoySkladError("Не задан MOYSKLAD_TOKEN.")
        self._token = token
        self._timeout = timeout
        self._opener = opener
        self._ssl_context = ssl.create_default_context()

    def _get_json(self, url: str, params: dict | None = None):
        if params:
            url = f"{url}?{urlencode(params, doseq=True)}"
        parsed = urlparse(url)
        if parsed.scheme != "https" or parsed.hostname not in {
            "api.moysklad.ru", "online.moysklad.ru"
        }:
            raise MoySkladError("API вернул небезопасный адрес продолжения выборки.")
        request = Request(
            url,
            headers={
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json;charset=utf-8",
                "Accept-Encoding": "gzip",
                "User-Agent": "UralVapeRegionsBot/1.0 (read-only)",
            },
            method="GET",
        )
        error: Exception | None = None
        for attempt in range(3):
            try:
                response = self._opener(request, timeout=self._timeout, context=self._ssl_context)
                with response:
                    body = response.read()
                    if str(response.headers.get("Content-Encoding", "")).casefold() == "gzip":
                        body = gzip.decompress(body)
                    return json.loads(body.decode("utf-8"))
            except HTTPError as current_error:
                error = current_error
                if current_error.code not in {429, 500, 502, 503, 504} or attempt == 2:
                    break
            except (TimeoutError, URLError) as current_error:
                error = current_error
                if attempt == 2:
                    break
            except Exception as current_error:
                error = current_error
                if "timed out" not in str(current_error).casefold() or attempt == 2:
                    break
            time.sleep(2 ** attempt)
        if isinstance(error, HTTPError):
            try:
                body = error.read()
                if str(error.headers.get("Content-Encoding", "")).casefold() == "gzip":
                    body = gzip.decompress(body)
                payload = json.loads(body.decode("utf-8"))
                errors = payload.get("errors", []) if isinstance(payload, dict) else []
                details = "; ".join(
                    str(item.get("error", "")).strip()
                    for item in errors
                    if isinstance(item, dict) and item.get("error")
                )
            except Exception:
                details = ""
            endpoint = urlparse(url).path
            message = f"МойСклад отклонил GET {endpoint}: HTTP {error.code}"
            if details:
                message += f" — {details}"
            raise MoySkladError(message, status=error.code) from error
        raise MoySkladError(
            f"Ошибка чтения API МоегоСклада после 3 попыток: {error}"
        ) from error

    def _rows(self, endpoint: str, params: dict | None = None) -> list[dict]:
        url = f"{API_ROOT}/{endpoint.lstrip('/')}"
        query = {"limit": 1000, **(params or {})}
        result: list[dict] = []
        while url:
            payload = self._get_json(url, query)
            query = None
            rows = payload.get("rows", [])
            if not isinstance(rows, list):
                raise MoySkladError("API МоегоСклада вернул ответ неизвестного формата.")
            result.extend(row for row in rows if isinstance(row, dict))
            url = payload.get("meta", {}).get("nextHref")
        return result

    def stores(self) -> list[dict]:
        return self._rows("entity/store")

    def current_availability(self, store_ids: tuple[str, ...]) -> list[dict]:
        """Return MoySklad's native «Доступно» value for selected stores."""
        endpoint = f"{API_ROOT}/report/stock/bystore/current"
        try:
            payload = self._get_json(
                endpoint,
                {
                    "stockType": "quantity",
                    "filter": f"storeId={','.join(store_ids)}",
                },
            )
        except MoySkladError as error:
            # Some accounts/API revisions reject the documented multi-store filter.
            # The unfiltered report is still read-only; selected stores are filtered
            # locally by _available_stock.
            if error.status != 400:
                raise
            payload = self._get_json(endpoint, {"stockType": "quantity"})
        if not isinstance(payload, list):
            raise MoySkladError("API МоегоСклада вернул неизвестный формат отчёта «Доступно».")
        return [row for row in payload if isinstance(row, dict)]

    def assortment(self) -> list[dict]:
        return self._rows("entity/assortment")

    def customer_order_print_templates(self) -> list[dict]:
        """Return embedded and custom print templates for customer orders."""
        result: list[dict] = []
        for kind in ("embeddedtemplate", "customtemplate"):
            result.extend(self._rows(f"entity/customerorder/metadata/{kind}"))
        return result

    def export_customer_order_pdf_url(self, order_id: str, template: dict) -> str:
        """Start a non-mutating print export and return MoySklad's temporary PDF URL."""
        order_id = order_id.strip()
        template_meta = template.get("meta") if isinstance(template, dict) else None
        if not order_id or not isinstance(template_meta, dict):
            raise MoySkladError("Не указан заказ или шаблон печатной формы.")
        body = json.dumps(
            {"template": {"meta": template_meta}, "extension": "pdf"},
            ensure_ascii=False,
        ).encode("utf-8")
        request = Request(
            f"{API_ROOT}/entity/customerorder/{order_id}/export",
            data=body,
            headers={
                "Authorization": f"Bearer {self._token}",
                "Accept": "application/json;charset=utf-8",
                "Content-Type": "application/json;charset=utf-8",
                "User-Agent": "UralVapeRegionsBot/1.0 (print-export)",
            },
            method="POST",
        )

        class NoRedirect(HTTPRedirectHandler):
            def redirect_request(self, req, fp, code, msg, headers, newurl):
                return None

        opener = build_opener(HTTPSHandler(context=self._ssl_context), NoRedirect())
        location = ""
        try:
            with opener.open(request, timeout=self._timeout) as response:
                location = str(response.headers.get("Location") or response.geturl() or "")
        except HTTPError as error:
            if error.code not in {301, 302, 303, 307, 308}:
                raise MoySkladError(
                    f"МойСклад не сформировал печатную форму заказа: HTTP {error.code}",
                    status=error.code,
                ) from error
            location = str(error.headers.get("Location") or "")
        parsed = urlparse(location)
        hostname = (parsed.hostname or "").casefold()
        if parsed.scheme != "https" or not (
            hostname == "moysklad.ru" or hostname.endswith(".moysklad.ru")
        ):
            raise MoySkladError("МойСклад не вернул безопасную ссылку на PDF.")
        return location

    def product_folders(self) -> list[dict]:
        return self._rows("entity/productfolder")

    def sales_channels(self) -> list[dict]:
        return self._rows("entity/saleschannel")

    def changed_customer_orders(self, updated_from: str) -> list[dict]:
        """Return recently changed customer orders without modifying MoySklad."""
        return self._rows(
            "entity/customerorder",
            {
                "limit": 100,
                "filter": f"updated>={updated_from}",
                "expand": "agent,state,salesChannel",
            },
        )


def build_product_folder_mapping(
    token: str,
    destination: Path,
    *,
    client: MoySkladClient | None = None,
    categories: dict[str, str] | None = None,
    category_choices: tuple[str, ...] = BONUS_CATEGORIES,
) -> int:
    """Export only terminal MoySklad folders for manual category mapping."""
    client = client or MoySkladClient(token)
    folders = client.product_folders()
    active_rows = []
    for folder in folders:
        if folder.get("archived") is True:
            continue
        name = str(folder.get("name", "")).strip()
        parent_path = str(folder.get("pathName", "")).strip(" /\t")
        if not name:
            continue
        full_path = f"{parent_path}/{name}" if parent_path else name
        active_rows.append((full_path, str(folder.get("id", "")).strip()))
    all_paths = {row[0].casefold() for row in active_rows}
    rows = [
        row for row in active_rows
        if not any(path.startswith(row[0].casefold() + "/") for path in all_paths)
    ]
    rows = sorted(set(rows), key=lambda row: row[0].casefold())
    if not rows:
        raise MoySkladError("В МоемСкладе не найдены активные папки товаров.")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Классификация папок"
    headers = [
        "Полный путь папки",
        "Категория",
        "Комментарий",
        "ID папки МоегоСклада",
    ]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2F5597")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for full_path, folder_id in rows:
        category = (categories or {}).get(full_path.casefold(), "")
        sheet.append([full_path, category, "", folder_id])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:D{sheet.max_row}"
    sheet.column_dimensions["A"].width = 85
    sheet.column_dimensions["B"].width = 32
    sheet.column_dimensions["C"].width = 45
    sheet.column_dimensions["D"].width = 40

    guide = workbook.create_sheet("Справочник")
    guide.append(["Допустимые категории"])
    guide["A1"].font = Font(bold=True, color="FFFFFF")
    guide["A1"].fill = PatternFill("solid", fgColor="2F5597")
    for category in category_choices:
        guide.append([category])
    guide.column_dimensions["A"].width = 35
    validation = DataValidation(
        type="list",
        # Keep spare rows in the source range so an administrator can add new
        # monthly categories without losing the dropdown in the main sheet.
        formula1=f"'Справочник'!$A$2:$A${MAX_BONUS_CATEGORY_CHOICES + 1}",
        allow_blank=True,
    )
    validation.error = "Выберите категорию из списка."
    validation.errorTitle = "Неизвестная категория"
    validation.prompt = "Выберите премиальную категорию или «Не учитывать»."
    validation.promptTitle = "Категория"
    sheet.add_data_validation(validation)
    validation.add(f"B2:B{sheet.max_row}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    return len(rows)


def _select_stores(rows: list[dict], wanted: tuple[str, ...]) -> tuple[dict, ...]:
    by_name: dict[str, list[dict]] = {}
    for row in rows:
        by_name.setdefault(str(row.get("name", "")).strip().casefold(), []).append(row)
    selected = []
    missing = []
    for name in wanted:
        matches = by_name.get(name.strip().casefold(), [])
        if not matches:
            missing.append(name)
        elif len(matches) > 1:
            raise MoySkladError(f"В МоемСкладе найдено несколько складов с названием «{name}».")
        else:
            selected.append(matches[0])
    if missing:
        raise MoySkladError("Не найдены склады: " + ", ".join(missing) + ".")
    return tuple(selected)


def _price_type_names(
    assortment: list[dict], cash_name: str, cashless_name: str
) -> tuple[str, str]:
    names = {
        str(price.get("priceType", {}).get("name", "")).strip()
        for item in assortment
        for price in item.get("salePrices", [])
        if price.get("priceType", {}).get("name")
    }

    def resolve(
        configured: str,
        marker: str,
        exclude: str = "",
        preferred: str = "",
    ) -> str:
        if configured:
            exact = [name for name in names if name.casefold() == configured.casefold()]
            if not exact:
                raise MoySkladError(f"Тип цены «{configured}» не найден в МоемСкладе.")
            return exact[0]
        preferred_match = [name for name in names if name.casefold() == preferred.casefold()]
        if preferred and preferred_match:
            return preferred_match[0]
        matches = [
            name for name in names
            if marker in name.casefold() and (not exclude or exclude not in name.casefold())
        ]
        if len(matches) != 1:
            available = ", ".join(sorted(names)) or "нет доступных типов цен"
            raise MoySkladError(
                f"Не удалось однозначно определить цену «{marker}». "
                f"Укажите её точное название в .env. Доступны: {available}."
            )
        return matches[0]

    cashless = resolve(
        cashless_name.strip(), "безнал", preferred="от 50т.р. безнал"
    )
    cash = resolve(
        cash_name.strip(), "нал", "безнал", preferred="от 50т.р. нал"
    )
    return cash, cashless


def _sale_price(item: dict, price_type: str) -> Decimal | None:
    for price in item.get("salePrices", []):
        if str(price.get("priceType", {}).get("name", "")).strip().casefold() == price_type.casefold():
            return _money(price.get("value"))
    return None


def _available_stock(rows: list[dict], store_ids: tuple[str, ...]) -> dict[str, tuple[Decimal, ...]]:
    wanted = {store_id: index for index, store_id in enumerate(store_ids)}
    result: dict[str, tuple[Decimal, ...]] = {}
    for row in rows:
        assortment_id = str(row.get("assortmentId", "")).strip()
        index = wanted.get(str(row.get("storeId", "")).strip())
        if not assortment_id or index is None:
            continue
        quantity = row.get("quantity")
        if quantity is None:
            continue
        values = list(result.get(assortment_id, (Decimal(0),) * len(store_ids)))
        values[index] = max(Decimal(0), Decimal(str(quantity)))
        result[assortment_id] = tuple(values)
    return {key: values for key, values in result.items() if sum(values) > 0}


def build_price_from_moysklad(
    token: str,
    destination: Path,
    *,
    store_names: tuple[str, ...] = DEFAULT_STORES,
    cash_price_type: str = "",
    cashless_price_type: str = "",
    client: MoySkladClient | None = None,
) -> MoySkladPriceResult:
    client = client or MoySkladClient(token)
    selected_stores = _select_stores(client.stores(), store_names)
    store_ids = tuple(
        _canonical_href(str(store.get("meta", {}).get("href", ""))).rsplit("/", 1)[-1]
        for store in selected_stores
    )
    if any(not store_id for store_id in store_ids):
        raise MoySkladError("Для выбранных складов не получены API-идентификаторы.")

    stock = _available_stock(client.current_availability(store_ids), store_ids)
    assortment = client.assortment()
    cash_name, cashless_name = _price_type_names(
        assortment, cash_price_type, cashless_price_type
    )
    groups: dict[str, list[tuple[str, str, Decimal, Decimal, tuple[Decimal, ...]]]] = {}
    for item in assortment:
        if item.get("archived") is True:
            continue
        entity_type = str(item.get("meta", {}).get("type", ""))
        if entity_type not in {"product", "variant"}:
            continue
        assortment_id = _canonical_href(
            str(item.get("meta", {}).get("href", ""))
        ).rsplit("/", 1)[-1]
        balances = stock.get(assortment_id)
        if not balances:
            continue
        cash = _sale_price(item, cash_name)
        cashless = _sale_price(item, cashless_name)
        name = str(item.get("name", "")).strip()
        if not name or cash is None or cashless is None:
            continue
        path_name = str(item.get("pathName", "")).strip(" /\t") or "Без группы"
        group_path = path_name if "/" in path_name else f"МойСклад/{path_name}"
        code = str(item.get("code") or item.get("article") or "").strip()
        groups.setdefault(group_path, []).append((code, name, cash, cashless, balances))

    if not groups:
        raise MoySkladError(
            "Не найдено доступных позиций с обеими ценами на выбранных складах."
        )

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Общий прайс"
    today = date.today()
    sheet.append(["Прайс актуален", today])
    sheet.append([])
    headers = ["Наименование", "от 50т.р. нал", "от 50т.р. безнал", *store_names]
    sheet.append(headers)
    header_fill = PatternFill("solid", fgColor="2F5597")
    group_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in sheet[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    item_count = 0
    for group_path in sorted(groups, key=str.casefold):
        sheet.append([group_path])
        for cell in sheet[sheet.max_row]:
            cell.fill = group_fill
            cell.font = Font(bold=True)
        for _, name, cash, cashless, balances in sorted(groups[group_path], key=lambda row: row[1].casefold()):
            sheet.append([name, float(cash), float(cashless), *(float(value) for value in balances)])
            for column in range(2, 4):
                sheet.cell(sheet.max_row, column).number_format = "0.00"
            for column in range(4, 4 + len(store_names)):
                sheet.cell(sheet.max_row, column).number_format = "0.###"
            item_count += 1
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{chr(67 + len(store_names))}{sheet.max_row}"
    sheet.column_dimensions["A"].width = 80
    sheet.column_dimensions["B"].width = 20
    sheet.column_dimensions["C"].width = 23
    for column in range(4, 4 + len(store_names)):
        sheet.column_dimensions[chr(64 + column)].width = 16
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    return MoySkladPriceResult(
        destination, item_count, len(groups), store_names, cash_name, cashless_name
    )
