"""Read-only MoySklad bonus reports and monthly folder classification."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from moysklad_price import BONUS_CATEGORIES, MoySkladClient, MoySkladError, _canonical_href


OTHER_CATEGORY = "Прочее"
UNCLASSIFIED_CATEGORY = "Не классифицировано"
EXCLUDED_CATEGORY = "Не учитывать"
MAX_BONUS_CATEGORIES = 20


@dataclass(frozen=True)
class BonusClassification:
    categories: tuple[str, ...]
    folders: dict[str, str]


@dataclass(frozen=True)
class BonusDocument:
    kind: str
    moment: str
    name: str
    agent: str
    state: str
    sales_channel: str
    total: Decimal
    paid: Decimal
    categories: dict[str, Decimal]


@dataclass(frozen=True)
class BonusReportResult:
    path: Path
    document_count: int
    shipment_count: int
    return_count: int
    unclassified_paths: tuple[str, ...]


def month_bounds(month: str) -> tuple[str, str]:
    try:
        year, number = (int(part) for part in month.split("-", 1))
        start = datetime(year, number, 1)
    except (TypeError, ValueError) as error:
        raise MoySkladError("Месяц должен быть указан в формате ГГГГ-ММ.") from error
    if number == 12:
        end = datetime(year + 1, 1, 1)
    else:
        end = datetime(year, number + 1, 1)
    return start.strftime("%Y-%m-%d 00:00:00"), end.strftime("%Y-%m-%d 00:00:00")


def month_title(month: str) -> str:
    year, number = (int(part) for part in month.split("-", 1))
    names = (
        "январь", "февраль", "март", "апрель", "май", "июнь",
        "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
    )
    return f"{names[number - 1]} {year}"


def _full_folder_path(folder: dict) -> str:
    name = str(folder.get("name", "")).strip()
    parent = str(folder.get("pathName", "")).strip(" /\t")
    return f"{parent}/{name}" if parent else name


def leaf_folders(folders: list[dict]) -> list[dict]:
    active = [folder for folder in folders if not folder.get("archived") and folder.get("name")]
    paths = {_full_folder_path(folder).casefold() for folder in active}
    result = []
    for folder in active:
        path = _full_folder_path(folder)
        prefix = path.casefold() + "/"
        if not any(other.startswith(prefix) for other in paths):
            result.append(folder)
    return sorted(result, key=lambda folder: _full_folder_path(folder).casefold())


def _validate_categories(values: list[str]) -> tuple[str, ...]:
    categories = []
    seen = set()
    for value in values:
        category = str(value or "").strip()
        if not category:
            continue
        key = category.casefold()
        if key == EXCLUDED_CATEGORY.casefold() and category != EXCLUDED_CATEGORY:
            raise MoySkladError(f"Служебную категорию нужно оставить точно как «{EXCLUDED_CATEGORY}».")
        if key in seen:
            raise MoySkladError(f"В справочнике категория «{category}» указана повторно.")
        if len(category) > 60:
            raise MoySkladError(f"Название категории «{category}» длиннее 60 символов.")
        seen.add(key)
        categories.append(category)
    if EXCLUDED_CATEGORY.casefold() not in seen:
        raise MoySkladError(f"В справочнике должна оставаться служебная категория «{EXCLUDED_CATEGORY}».")
    if len(categories) < 2:
        raise MoySkladError("Добавьте в справочник хотя бы одну премиальную категорию.")
    if len(categories) > MAX_BONUS_CATEGORIES + 1:
        raise MoySkladError(f"Допускается не более {MAX_BONUS_CATEGORIES} премиальных категорий.")
    return tuple(categories)


def read_classification(path: Path) -> BonusClassification:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.worksheets[0]
    if len(workbook.worksheets) > 1:
        category_values = [
            str(row[0].value or "").strip()
            for row in workbook.worksheets[1].iter_rows(
                min_row=2, min_col=1, max_col=1
            )
        ]
        categories = _validate_categories(category_values)
    else:
        categories = tuple(BONUS_CATEGORIES)
    headers = {
        str(cell.value).strip().casefold(): index
        for index, cell in enumerate(sheet[1])
        if cell.value is not None
    }
    path_index = headers.get("полный путь папки")
    category_index = headers.get("категория")
    if path_index is None or category_index is None:
        workbook.close()
        raise MoySkladError("В таблице нужны колонки «Полный путь папки» и «Категория».")
    entered: dict[str, str] = {}
    errors = []
    allowed = set(categories)
    for row_number, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
        folder_path = str(values[path_index] or "").strip(" /\t")
        category = str(values[category_index] or "").strip()
        if not folder_path:
            continue
        if not category:
            errors.append(f"строка {row_number}: не выбрана категория")
            continue
        if category not in allowed:
            errors.append(f"строка {row_number}: неизвестная категория «{category}»")
            continue
        key = folder_path.casefold()
        if key in entered and entered[key] != category:
            errors.append(f"строка {row_number}: папка указана повторно с другой категорией")
        entered[key] = category
    workbook.close()
    if errors:
        details = "; ".join(errors[:8])
        if len(errors) > 8:
            details += f"; и ещё {len(errors) - 8}"
        raise MoySkladError("Таблица классификации не прошла проверку: " + details + ".")
    if not entered:
        raise MoySkladError("В таблице нет заполненных товарных папок.")
    # Older exports contained the whole tree. Keep only terminal folders so the
    # saved monthly rules have the same semantics as new leaf-only templates.
    mapping = {
        key: category
        for key, category in entered.items()
        if not any(other.startswith(key + "/") for other in entered)
    }
    return BonusClassification(categories=categories, folders=mapping)


def save_classification(source: Path, destination: Path) -> int:
    classification = read_classification(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 2,
        "categories": list(classification.categories),
        "folders": classification.folders,
    }
    temporary = destination.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(destination)
    return len(classification.folders)


def load_classification(path: Path) -> BonusClassification:
    if not path.exists():
        raise MoySkladError("Для выбранного месяца ещё не загружена классификация товарных папок.")
    payload = json.loads(path.read_text(encoding="utf-8"))
    folders = payload.get("folders", {})
    if not isinstance(folders, dict) or not folders:
        raise MoySkladError("Сохранённая классификация товарных папок повреждена или пуста.")
    raw_categories = payload.get("categories")
    if isinstance(raw_categories, list):
        categories = _validate_categories([str(value) for value in raw_categories])
    else:
        # Version 1 stored only folder mappings and used the original fixed list.
        categories = _validate_categories(list(BONUS_CATEGORIES))
    normalized = {str(key).casefold(): str(value) for key, value in folders.items()}
    unknown = sorted(set(normalized.values()) - set(categories))
    if unknown:
        raise MoySkladError("В сохранённой классификации неизвестные категории: " + ", ".join(unknown) + ".")
    return BonusClassification(categories=categories, folders=normalized)


def _expanded_rows(client: MoySkladClient, endpoint: str, start: str, end: str, expand: str) -> list[dict]:
    return client._rows(
        endpoint,
        {
            "limit": 100,
            "filter": f"moment>={start};moment<{end};applicable=true",
            "expand": expand,
        },
    )


def fetch_month_documents(client: MoySkladClient, month: str) -> tuple[list[dict], list[dict]]:
    start, end = month_bounds(month)
    demands = _expanded_rows(client, "entity/demand", start, end, "agent,state,salesChannel,positions")
    returns = _expanded_rows(client, "entity/salesreturn", start, end, "agent,state,demand.salesChannel,positions")
    return demands, returns


def _name(value: object, default: str = "—") -> str:
    return str(value.get("name", default)).strip() if isinstance(value, dict) else default


def _channel(document: dict) -> str:
    channel = _name(document.get("salesChannel"), "")
    if channel:
        return channel
    demand = document.get("demand")
    return _name(demand.get("salesChannel"), "") if isinstance(demand, dict) else ""


def sales_channels(documents: tuple[list[dict], list[dict]]) -> tuple[str, ...]:
    values = {_channel(document) for rows in documents for document in rows}
    return tuple(sorted((value for value in values if value), key=str.casefold))


def _minor_money(value: object) -> Decimal:
    try:
        return Decimal(str(value or 0)) / Decimal(100)
    except Exception:
        return Decimal(0)


def _position_sum(position: dict) -> Decimal:
    quantity = Decimal(str(position.get("quantity") or 0))
    price = _minor_money(position.get("price"))
    discount = Decimal(str(position.get("discount") or 0))
    return price * quantity * (Decimal(1) - discount / Decimal(100))


def _assortment_paths(client: MoySkladClient) -> dict[str, str]:
    result = {}
    for item in client.assortment():
        href = _canonical_href(str(item.get("meta", {}).get("href", "")))
        path = str(item.get("pathName", "")).strip(" /\t")
        if href:
            result[href] = path
    return result


def _document_positions(document: dict, client: MoySkladClient) -> list[dict]:
    positions = document.get("positions", {})
    if isinstance(positions, dict) and isinstance(positions.get("rows"), list):
        rows = positions["rows"]
        meta = positions.get("meta", {})
        if int(meta.get("size", len(rows)) or 0) <= len(rows):
            return [row for row in rows if isinstance(row, dict)]
        href = meta.get("href")
    else:
        href = positions.get("meta", {}).get("href") if isinstance(positions, dict) else None
    if not href:
        return []
    payload = client._get_json(str(href), {"limit": 1000})
    return [row for row in payload.get("rows", []) if isinstance(row, dict)]


def prepare_documents(
    client: MoySkladClient,
    raw: tuple[list[dict], list[dict]],
    classification: BonusClassification,
    channel: str,
) -> tuple[list[BonusDocument], set[str]]:
    assortment_paths = _assortment_paths(client)
    result: list[BonusDocument] = []
    unknown: set[str] = set()
    report_categories = tuple(
        category for category in classification.categories
        if category != EXCLUDED_CATEGORY
    )
    for kind, rows, sign in (("Отгрузка", raw[0], Decimal(1)), ("Возврат", raw[1], Decimal(-1))):
        for document in rows:
            document_channel = _channel(document)
            if document_channel.casefold() != channel.casefold():
                continue
            amounts = {category: Decimal(0) for category in report_categories}
            amounts[OTHER_CATEGORY] = Decimal(0)
            amounts[UNCLASSIFIED_CATEGORY] = Decimal(0)
            for position in _document_positions(document, client):
                assortment = position.get("assortment", {})
                href = _canonical_href(str(assortment.get("meta", {}).get("href", "")))
                folder_path = assortment_paths.get(href, "")
                category = classification.folders.get(folder_path.casefold()) if folder_path else None
                value = _position_sum(position) * sign
                if category in amounts:
                    amounts[category] += value
                elif category == EXCLUDED_CATEGORY:
                    amounts[OTHER_CATEGORY] += value
                else:
                    amounts[UNCLASSIFIED_CATEGORY] += value
                    unknown.add(folder_path or "Без папки")
            result.append(BonusDocument(
                kind=kind,
                moment=str(document.get("moment", "")),
                name=str(document.get("name", "")),
                agent=_name(document.get("agent")),
                state=_name(document.get("state")),
                sales_channel=document_channel,
                total=_minor_money(document.get("sum")) * sign,
                paid=_minor_money(document.get("payedSum")) * sign,
                categories=amounts,
            ))
    result.sort(key=lambda row: (row.moment, row.kind, row.name))
    return result, unknown


def build_bonus_report(
    token: str,
    month: str,
    channel: str,
    classification_path: Path,
    destination: Path,
    *,
    client: MoySkladClient | None = None,
    raw_documents: tuple[list[dict], list[dict]] | None = None,
) -> BonusReportResult:
    client = client or MoySkladClient(token)
    classification = load_classification(classification_path)
    raw = raw_documents or fetch_month_documents(client, month)
    documents, unknown = prepare_documents(client, raw, classification, channel)
    if not documents:
        raise MoySkladError(f"За {month_title(month)} по каналу продаж «{channel}» проведённых документов не найдено.")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Отчёт"
    sheet.append(["Отчёт по премиальным категориям", month_title(month), channel])
    sheet.append([])
    category_headers = [
        category for category in classification.categories
        if category != EXCLUDED_CATEGORY
    ]
    headers = [
        "Тип", "Дата", "Документ", "Контрагент", "Статус", "Канал продаж",
        "Сумма документа", "Оплачено", *category_headers, OTHER_CATEGORY,
        UNCLASSIFIED_CATEGORY, "Контрольное расхождение",
    ]
    sheet.append(headers)
    blue = PatternFill("solid", fgColor="2F5597")
    return_fill = PatternFill("solid", fgColor="FCE4D6")
    total_fill = PatternFill("solid", fgColor="D9EAD3")
    for cell in sheet[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = blue
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for document in documents:
        categorized = sum(document.categories.values(), Decimal(0))
        sheet.append([
            document.kind,
            document.moment[:10],
            document.name,
            document.agent,
            document.state,
            document.sales_channel,
            float(document.total),
            float(document.paid),
            *(float(document.categories[category]) for category in category_headers),
            float(document.categories[OTHER_CATEGORY]),
            float(document.categories[UNCLASSIFIED_CATEGORY]),
            float(document.total - categorized),
        ])
        if document.kind == "Возврат":
            for cell in sheet[sheet.max_row]:
                cell.fill = return_fill
    first_data_row = 4
    total_row = sheet.max_row + 1
    sheet.cell(total_row, 1, "ИТОГО")
    for column in range(7, len(headers) + 1):
        letter = sheet.cell(1, column).column_letter
        sheet.cell(total_row, column, f"=SUM({letter}{first_data_row}:{letter}{total_row - 1})")
    for cell in sheet[total_row]:
        cell.font = Font(bold=True)
        cell.fill = total_fill
    for row in range(first_data_row, total_row + 1):
        for column in range(7, len(headers) + 1):
            sheet.cell(row, column).number_format = '#,##0.00'
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{sheet.cell(3, len(headers)).column_letter}{total_row - 1}"
    fixed_widths = [13, 13, 18, 38, 22, 24, 19, 16]
    widths = fixed_widths + [24] * len(category_headers) + [15, 23, 23]
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[sheet.cell(1, index).column_letter].width = width

    if unknown:
        unknown_sheet = workbook.create_sheet("Не классифицировано")
        unknown_sheet.append(["Папка товара"])
        for folder_path in sorted(unknown, key=str.casefold):
            unknown_sheet.append([folder_path])
        unknown_sheet.column_dimensions["A"].width = 90
        unknown_sheet.freeze_panes = "A2"
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    workbook.close()
    return BonusReportResult(
        path=destination,
        document_count=len(documents),
        shipment_count=sum(row.kind == "Отгрузка" for row in documents),
        return_count=sum(row.kind == "Возврат" for row in documents),
        unclassified_paths=tuple(sorted(unknown, key=str.casefold)),
    )
