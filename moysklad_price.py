"""Read-only MoySklad integration used to build the common price list."""

from __future__ import annotations

import json
import gzip
import ssl
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Callable
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill


API_ROOT = "https://api.moysklad.ru/api/remap/1.2"
DEFAULT_STORES = ("Мордор", "Годзибасы", "Жможики")


class MoySkladError(RuntimeError):
    pass


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
        timeout: int = 45,
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
                "Accept": "application/json",
                "Accept-Encoding": "gzip",
                "User-Agent": "UralVapeRegionsBot/1.0 (read-only)",
            },
            method="GET",
        )
        try:
            response = self._opener(request, timeout=self._timeout, context=self._ssl_context)
            with response:
                body = response.read()
                if str(response.headers.get("Content-Encoding", "")).casefold() == "gzip":
                    body = gzip.decompress(body)
                return json.loads(body.decode("utf-8"))
        except MoySkladError:
            raise
        except Exception as error:
            raise MoySkladError(f"Ошибка чтения API МоегоСклада: {error}") from error

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
        payload = self._get_json(
            f"{API_ROOT}/report/stock/bystore/current",
            {
                "stockType": "quantity",
                "filter": f"storeId={','.join(store_ids)}",
            },
        )
        if not isinstance(payload, list):
            raise MoySkladError("API МоегоСклада вернул неизвестный формат отчёта «Доступно».")
        return [row for row in payload if isinstance(row, dict)]

    def assortment(self) -> list[dict]:
        return self._rows("entity/assortment")


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

    def resolve(configured: str, marker: str, exclude: str = "") -> str:
        if configured:
            exact = [name for name in names if name.casefold() == configured.casefold()]
            if not exact:
                raise MoySkladError(f"Тип цены «{configured}» не найден в МоемСкладе.")
            return exact[0]
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

    cashless = resolve(cashless_name.strip(), "безнал")
    cash = resolve(cash_name.strip(), "нал", "безнал")
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
