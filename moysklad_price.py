"""Read-only MoySklad integration used to build the common price list."""

from __future__ import annotations

import json
import gzip
import ssl
from copy import copy
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Callable
from urllib.parse import urlencode, urlparse
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


API_ROOT = "https://api.moysklad.ru/api/remap/1.2"
DEFAULT_STORES = ("Мордор", "Годзибасы", "Жможики")


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


def _header_text(value: object) -> str:
    return " ".join(
        str(value or "").casefold().replace("ё", "е").replace(".", " ").split()
    )


def _template_layout(sheet) -> dict:
    for row in range(1, min(sheet.max_row, 40) + 1):
        headers = [_header_text(sheet.cell(row, column).value) for column in range(1, sheet.max_column + 1)]
        name_column = next((i for i, value in enumerate(headers, 1) if value == "наименование"), None)
        cashless_column = next((i for i, value in enumerate(headers, 1) if "безнал" in value), None)
        cash_column = next(
            (i for i, value in enumerate(headers, 1) if "нал" in value and "безнал" not in value),
            None,
        )
        if name_column and cash_column and cashless_column:
            header_row = row
            code_column = next((i for i, value in enumerate(headers, 1) if value == "код"), 1)
            unit_column = next((i for i, value in enumerate(headers, 1) if value in {"ед изм", "единица измерения"}), None)
            break
    else:
        raise MoySkladError(
            "В шаблоне не найдены колонки «Наименование», «нал» и «безнал»."
        )

    group_row = None
    action_row = None
    product_row = None
    for row in range(header_row + 1, sheet.max_row + 1):
        code = str(sheet.cell(row, code_column).value or "").strip()
        name = str(sheet.cell(row, name_column).value or "").strip()
        cash = sheet.cell(row, cash_column).value
        cashless = sheet.cell(row, cashless_column).value
        if group_row is None and "/" in code and not name and cash is None and cashless is None:
            group_row = row
        if name and cash is not None and cashless is not None:
            if name.casefold().startswith("акция ") and action_row is None:
                action_row = row
            elif product_row is None:
                product_row = row
        if group_row and product_row and action_row:
            break
    product_row = product_row or action_row
    action_row = action_row or product_row
    if not group_row or not product_row:
        raise MoySkladError(
            "В шаблоне не найдены образцы строки товарной группы и товарной позиции."
        )
    return {
        "header_row": header_row,
        "code_column": code_column,
        "name_column": name_column,
        "unit_column": unit_column,
        "cash_column": cash_column,
        "cashless_column": cashless_column,
        "group_row": group_row,
        "product_row": product_row,
        "action_row": action_row,
    }


def validate_price_template(path: Path) -> dict:
    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as error:
        raise MoySkladError(f"Не удалось открыть Excel-шаблон: {error}") from error
    try:
        layout = _template_layout(workbook.active)
        return {
            "sheet_name": workbook.active.title,
            "header_row": layout["header_row"],
            "columns": workbook.active.max_column,
        }
    finally:
        workbook.close()


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
                "Accept": "application/json;charset=utf-8",
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
        except HTTPError as error:
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


def _row_blueprint(sheet, source_row: int) -> dict:
    merges = [
        (merged.min_col, merged.max_col)
        for merged in sheet.merged_cells.ranges
        if merged.min_row == source_row and merged.max_row == source_row
    ]
    return {
        "source_row": source_row,
        "height": sheet.row_dimensions[source_row].height,
        "cells": [
            (cell.value, copy(cell._style), cell.number_format)
            for cell in sheet[source_row]
        ],
        "merges": merges,
    }


def _apply_blueprint(sheet, target_row: int, blueprint: dict, *, formulas: bool) -> None:
    source_row = blueprint["source_row"]
    sheet.row_dimensions[target_row].height = blueprint["height"]
    for column, (value, style, number_format) in enumerate(blueprint["cells"], 1):
        cell = sheet.cell(target_row, column)
        cell._style = copy(style)
        cell.number_format = number_format
        if formulas and isinstance(value, str) and value.startswith("="):
            try:
                value = Translator(
                    value, origin=f"{get_column_letter(column)}{source_row}"
                ).translate_formula(f"{get_column_letter(column)}{target_row}")
            except Exception:
                pass
            cell.value = value
        else:
            cell.value = None
    for start_column, end_column in blueprint["merges"]:
        sheet.merge_cells(
            start_row=target_row,
            start_column=start_column,
            end_row=target_row,
            end_column=end_column,
        )


def _write_template_price(
    template_path: Path,
    destination: Path,
    groups: dict[str, list[tuple[str, str, Decimal, Decimal, tuple[Decimal, ...]]]],
) -> None:
    workbook = load_workbook(template_path, data_only=False)
    sheet = workbook.active
    layout = _template_layout(sheet)
    header_row = layout["header_row"]
    original_max_row = sheet.max_row
    original_max_column = sheet.max_column
    group_blueprint = _row_blueprint(sheet, layout["group_row"])
    product_blueprint = _row_blueprint(sheet, layout["product_row"])
    action_blueprint = _row_blueprint(sheet, layout["action_row"])

    for merged in list(sheet.merged_cells.ranges):
        if merged.max_row > header_row:
            sheet.unmerge_cells(str(merged))
    if original_max_row > header_row:
        sheet.delete_rows(header_row + 1, original_max_row - header_row)

    target_row = header_row + 1
    for group_path in sorted(groups, key=str.casefold):
        _apply_blueprint(sheet, target_row, group_blueprint, formulas=False)
        sheet.cell(target_row, layout["code_column"]).value = group_path
        target_row += 1
        for code, name, cash, cashless, _ in sorted(
            groups[group_path], key=lambda item: item[1].casefold()
        ):
            blueprint = (
                action_blueprint
                if name.casefold().startswith("акция ")
                else product_blueprint
            )
            _apply_blueprint(sheet, target_row, blueprint, formulas=True)
            sheet.cell(target_row, layout["code_column"]).value = code
            sheet.cell(target_row, layout["name_column"]).value = name
            if layout["unit_column"]:
                sheet.cell(target_row, layout["unit_column"]).value = "шт"
            sheet.cell(target_row, layout["cash_column"]).value = float(cash)
            sheet.cell(target_row, layout["cashless_column"]).value = float(cashless)
            target_row += 1

    for row in range(1, min(header_row, 40) + 1):
        for column in range(1, original_max_column + 1):
            if "актуален" in _header_text(sheet.cell(row, column).value):
                sheet.cell(row, min(column + 1, original_max_column)).value = date.today()
                break

    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()


def build_price_from_moysklad(
    token: str,
    destination: Path,
    *,
    store_names: tuple[str, ...] = DEFAULT_STORES,
    cash_price_type: str = "",
    cashless_price_type: str = "",
    template_path: Path | None = None,
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

    if template_path is not None:
        if not template_path.is_file():
            raise MoySkladError("Шаблон прайса не загружен.")
        _write_template_price(template_path, destination, groups)
        return MoySkladPriceResult(
            destination,
            sum(len(items) for items in groups.values()),
            len(groups),
            store_names,
            cash_name,
            cashless_name,
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
