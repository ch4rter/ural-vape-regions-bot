import asyncio
import html
import logging
import os
import re
import shutil
import sqlite3
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from itertools import chain
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, FSInputFile, InlineKeyboardButton, InlineKeyboardMarkup, Message
from dotenv import load_dotenv
from openpyxl import load_workbook

from materials_db import Material, MaterialsDB
from prices_db import WAREHOUSES, GroupDetails, PricesDB, parse_price_file, save_price_source


BASE_DIR = Path(__file__).resolve().parent
RESULTS_PER_PAGE = 10
PRICE_VARIANTS_PER_PAGE = 12
router = Router()
catalog: "Catalog"
materials_db: MaterialsDB
prices_db: PricesDB
admin_ids: set[int] = set()
active_excel_path: Path
managed_excel_path: Path
price_storage_path: Path


class AppState(StatesGroup):
    region_search = State()
    price_search = State()


class AdminState(StatesGroup):
    product_name = State()
    product_rename = State()
    section_name = State()
    section_rename = State()
    material_upload = State()
    excel_upload = State()
    excel_confirmation = State()
    price_upload = State()
    price_confirmation = State()


def normalize(value: str) -> str:
    value = value.casefold().replace("ё", "е")
    value = re.sub(r"[\s\-–—_,.;:()]+", " ", value)
    return value.strip()


@dataclass(frozen=True)
class Entry:
    name: str
    manager: str
    location: str = ""


class Catalog:
    def __init__(self, path: Path):
        self.path = path
        self.entries: list[Entry] = []
        self.by_name: dict[str, list[int]] = {}
        self.reload()

    def reload(self) -> None:
        if not self.path.exists():
            raise FileNotFoundError(f"Не найден файл {self.path.name}.")
        workbook = load_workbook(self.path, read_only=True, data_only=True)
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        first_row = next(rows, None)
        if first_row is None:
            workbook.close()
            raise ValueError("Excel-файл пуст.")

        headers = [normalize(str(cell or "")) for cell in first_row]
        manager_idx = next((i for i, value in enumerate(headers) if "менеджер" in value), None)
        territory_idx = next((i for i, value in enumerate(headers) if "территор" in value), None)
        name_idx = next(
            (i for i, value in enumerate(headers) if "назван" in value or "регион" in value), None
        )
        location_idx = next((i for i, value in enumerate(headers) if "местополож" in value), None)
        has_header = manager_idx is not None and (territory_idx is not None or name_idx is not None)
        if has_header:
            name_idx = territory_idx if territory_idx is not None else name_idx
            data_rows = rows
            start_row = 2
        else:
            name_idx, manager_idx, location_idx = 0, 1, None
            data_rows = chain([first_row], rows)
            start_row = 1

        entries: list[Entry] = []
        by_name: dict[str, list[int]] = {}
        for row_number, row in enumerate(data_rows, start=start_row):
            name = str(row[name_idx]).strip() if len(row) > name_idx and row[name_idx] is not None else ""
            manager = (
                str(row[manager_idx]).strip()
                if len(row) > manager_idx and row[manager_idx] is not None else ""
            )
            location = (
                str(row[location_idx]).strip()
                if location_idx is not None and len(row) > location_idx and row[location_idx] is not None
                else ""
            )
            if not name and not manager:
                continue
            if not name or not manager:
                logging.warning("Пропущена неполная строка Excel: %s", row_number)
                continue
            index = len(entries)
            entries.append(Entry(name, manager, location))
            by_name.setdefault(normalize(name), []).append(index)
        workbook.close()
        if not entries:
            raise ValueError("В Excel нет заполненных записей.")
        self.entries, self.by_name = entries, by_name
        logging.info("Загружено строк из Excel: %s", len(entries))

    def exact(self, query: str) -> list[Entry]:
        return [self.entries[i] for i in self.by_name.get(normalize(query), [])]

    def suggestions(self, query: str, limit: int = 5) -> list[int]:
        query_norm = normalize(query)
        if not query_norm:
            return []
        ranked = []
        for name_norm, indexes in self.by_name.items():
            ratio = SequenceMatcher(None, query_norm, name_norm).ratio()
            contains = query_norm in name_norm or name_norm in query_norm
            if ratio >= 0.58 or contains:
                ranked.append((ratio + (0.15 if contains else 0), indexes[0]))
        ranked.sort(reverse=True)
        return [index for _, index in ranked[:limit]]


def validate_excel(path: Path) -> tuple[Catalog, tuple[str, ...]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    first_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    sheet_name = sheet.title
    workbook.close()
    if not first_row:
        raise ValueError("Excel-файл пуст.")
    headers = tuple(str(value or "").strip() for value in first_row)
    normalized = [normalize(value) for value in headers]
    required = {
        "местоположение": any("местополож" in value for value in normalized),
        "территория": any("территор" in value for value in normalized),
        "менеджер": any("менеджер" in value for value in normalized),
    }
    missing = [name for name, present in required.items() if not present]
    if missing:
        raise ValueError(f"Не найдены обязательные колонки: {', '.join(missing)}.")
    return Catalog(path), (sheet_name, *headers)


def is_admin(user_id: int | None) -> bool:
    return user_id is not None and user_id in admin_ids


def main_menu(user_id: int | None) -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton(text="🔎 Поиск менеджера", callback_data="main:region")],
        [InlineKeyboardButton(text="💰 Цены и наличие", callback_data="main:prices")],
        [InlineKeyboardButton(text="🗃 База данных", callback_data="main:database")],
    ]
    if is_admin(user_id):
        rows.append([InlineKeyboardButton(text="⚙️ Управление базой", callback_data="main:admin")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def back_main() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[[InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")]]
    )


async def show_main(target: Message, user_id: int | None, *, edit: bool = False) -> None:
    text = (
        "👋 <b>Добро пожаловать!</b>\n\n"
        "Здесь можно найти ответственного менеджера по территории или получить "
        "рабочие материалы по продукции.\n\n"
        "Выберите нужный раздел:"
    )
    if edit:
        await target.edit_text(text, reply_markup=main_menu(user_id))
    else:
        await target.answer(text, reply_markup=main_menu(user_id))


@router.message(CommandStart())
@router.message(Command("menu"))
async def command_menu(message: Message, state: FSMContext) -> None:
    await cleanup_pending_excel(state)
    await state.clear()
    await show_main(message, message.from_user.id if message.from_user else None)


@router.message(Command("cancel"))
async def cancel(message: Message, state: FSMContext) -> None:
    await cleanup_pending_excel(state)
    await state.clear()
    await message.answer("✅ Текущее действие отменено.")
    await show_main(message, message.from_user.id if message.from_user else None)


@router.callback_query(F.data == "main:menu")
async def callback_menu(callback: CallbackQuery, state: FSMContext) -> None:
    await cleanup_pending_excel(state)
    await state.clear()
    await show_main(callback.message, callback.from_user.id, edit=True)
    await callback.answer()


@router.callback_query(F.data == "main:region")
async def open_region_search(callback: CallbackQuery, state: FSMContext) -> None:
    await state.set_state(AppState.region_search)
    await callback.message.edit_text(
        "🔎 <b>Поиск менеджера</b>\n\n"
        "Отправьте название города, посёлка, деревни, области или другой территории.\n\n"
        "Например: <code>Питер</code>",
        reply_markup=back_main(),
    )
    await callback.answer()


def unique_results(entries: list[Entry]) -> list[tuple[str, str, str]]:
    return list(dict.fromkeys((entry.name, entry.location, entry.manager) for entry in entries))


def format_result(entries: list[Entry], page: int = 0) -> str:
    unique = unique_results(entries)
    if len(unique) == 1:
        name, location, manager = unique[0]
        lines = ["✅ <b>Менеджер найден</b>", "", f"📍 Территория: <b>{html.escape(name)}</b>"]
        if location:
            lines.append(f"🗺 Местоположение: <b>{html.escape(location)}</b>")
        lines.append(f"👤 Менеджер: <b>{html.escape(manager)}</b>")
        return "\n".join(lines)
    page_count = (len(unique) + RESULTS_PER_PAGE - 1) // RESULTS_PER_PAGE
    page = max(0, min(page, page_count - 1))
    start = page * RESULTS_PER_PAGE
    lines = ["🔎 <b>Найдено несколько вариантов</b>", f"Всего совпадений: <b>{len(unique)}</b>"]
    if page_count > 1:
        lines.append(f"Страница <b>{page + 1}</b> из <b>{page_count}</b>")
    for number, (name, location, manager) in enumerate(
        unique[start : start + RESULTS_PER_PAGE], start=start + 1
    ):
        block = [f"\n<b>{number}.</b> 📍 Территория: <b>{html.escape(name)}</b>"]
        if location:
            block.append(f"🗺 Местоположение: <b>{html.escape(location)}</b>")
        block.append(f"👤 Менеджер: <b>{html.escape(manager)}</b>")
        lines.append("\n".join(block))
    return "\n".join(lines)


def suggestion_label(entry: Entry, variant_count: int = 1) -> str:
    label = (
        f"{entry.name} — {variant_count} вариантов"
        if variant_count > 1
        else (f"{entry.name} — {entry.location}" if entry.location else entry.name)
    )
    return label if len(label) <= 64 else f"{label[:61]}..."


def pagination_keyboard(entry_index: int, entries: list[Entry], page: int) -> InlineKeyboardMarkup:
    total = len(unique_results(entries))
    page_count = (total + RESULTS_PER_PAGE - 1) // RESULTS_PER_PAGE
    buttons = []
    if page > 0:
        buttons.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"page:{entry_index}:{page - 1}"))
    if page + 1 < page_count:
        buttons.append(InlineKeyboardButton(text="Далее ➡️", callback_data=f"page:{entry_index}:{page + 1}"))
    rows = [buttons] if buttons else []
    rows.append([InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.message(AppState.region_search, F.text)
async def search_region(message: Message) -> None:
    query = message.text.strip()
    if len(query) > 200:
        await message.answer(
            "⚠️ Запрос получился слишком длинным. Отправьте только название территории.",
            reply_markup=back_main(),
        )
        return
    found = catalog.exact(query)
    if found:
        index = catalog.by_name[normalize(query)][0]
        await message.answer(format_result(found), reply_markup=pagination_keyboard(index, found, 0))
        return
    indexes = catalog.suggestions(query)
    if not indexes:
        await message.answer(
            "🤷 <b>Ничего не найдено</b>\n\n"
            "Проверьте написание или попробуйте указать другое название территории.",
            reply_markup=back_main(),
        )
        return
    rows = []
    for index in indexes:
        entry = catalog.entries[index]
        count = len(catalog.by_name[normalize(entry.name)])
        rows.append([InlineKeyboardButton(text=suggestion_label(entry, count), callback_data=f"place:{index}")])
    rows.append([InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")])
    await message.answer(
        "Точного совпадения не найдено. Возможно, вы имели в виду один из вариантов:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )


@router.message(AppState.region_search)
async def region_non_text(message: Message) -> None:
    await message.answer(
        "Пожалуйста, отправьте название территории обычным текстовым сообщением.",
        reply_markup=back_main(),
    )


@router.callback_query(F.data.startswith("place:"))
async def select_suggestion(callback: CallbackQuery) -> None:
    try:
        index = int(callback.data.split(":", 1)[1])
        selected = catalog.entries[index]
    except (ValueError, IndexError):
        await callback.answer("Выполните поиск ещё раз.", show_alert=True)
        return
    found = catalog.exact(selected.name)
    await callback.message.edit_text(format_result(found), reply_markup=pagination_keyboard(index, found, 0))
    await callback.answer()


@router.callback_query(F.data.startswith("page:"))
async def change_page(callback: CallbackQuery) -> None:
    try:
        _, raw_index, raw_page = callback.data.split(":", 2)
        index, page = int(raw_index), int(raw_page)
        selected = catalog.entries[index]
    except (ValueError, IndexError):
        await callback.answer("Выполните поиск ещё раз.", show_alert=True)
        return
    found = catalog.exact(selected.name)
    await callback.message.edit_text(
        format_result(found, page), reply_markup=pagination_keyboard(index, found, page)
    )
    await callback.answer()


def money(value: Decimal) -> str:
    quantized = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{quantized:,.2f}".replace(",", " ").replace(".", ",")


def discounted(value: Decimal, percent: int) -> Decimal:
    return value * (Decimal(100 - percent) / Decimal(100))


def variant_word(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "вариант"
    if count % 10 in (2, 3, 4) and count % 100 not in (12, 13, 14):
        return "варианта"
    return "вариантов"


def price_group_label(display_name: str, category_name: str) -> str:
    label = f"{display_name} · {category_name}"
    return label if len(label) <= 64 else f"{label[:61]}..."


def format_price_group(details: GroupDetails) -> str:
    lines = [
        f"💰 <b>{html.escape(details.summary.display_name)}</b>",
        f"Категория: <b>{html.escape(details.summary.category_name)}</b>",
        f"Уникальных вариантов: <b>{details.unique_variants}</b>",
        "",
        "🏢 <b>Наличие и ассортимент</b>",
    ]
    for warehouse in ("center", "west", "ural"):
        count = details.summary.warehouse_counts.get(warehouse, 0)
        if count:
            lines.append(f"✅ {WAREHOUSES[warehouse]} — <b>{count}</b> {variant_word(count)}")
        else:
            lines.append(f"❌ {WAREHOUSES[warehouse]} — нет в прайсе")
    lines.extend(["", "💳 <b>Цены внутри группы</b>"])
    for number, tier in enumerate(details.tiers, 1):
        if len(details.tiers) > 1:
            lines.extend([
                "",
                f"<b>Ценовой уровень {number}</b> · {tier.variant_count} {variant_word(tier.variant_count)}",
            ])
        for percent in (0, 5, 10, 15):
            cash = money(discounted(tier.cash, percent))
            cashless = money(discounted(tier.cashless, percent))
            label = "Без скидки" if percent == 0 else f"Скидка {percent}%"
            lines.extend([
                "",
                f"<b>{label}</b>",
                f"• Нал — <b>{cash} ₽</b>",
                f"• Безнал — <b>{cashless} ₽</b>",
            ])
    return "\n".join(lines)


def format_price_variants(summary, variants, page: int) -> str:
    page_count = max(1, (len(variants) + PRICE_VARIANTS_PER_PAGE - 1) // PRICE_VARIANTS_PER_PAGE)
    page = max(0, min(page, page_count - 1))
    start = page * PRICE_VARIANTS_PER_PAGE
    shown = variants[start : start + PRICE_VARIANTS_PER_PAGE]
    lines = [
        f"🎨 <b>{html.escape(summary.display_name)}</b>",
        f"Вкусов и цветов: <b>{len(variants)}</b>",
    ]
    if page_count > 1:
        lines.append(f"Страница <b>{page + 1}</b> из <b>{page_count}</b>")
    for number, variant in enumerate(shown, start=start + 1):
        warehouses = " · ".join(
            WAREHOUSES[key] for key in ("center", "west", "ural") if key in variant.warehouses
        )
        lines.extend([
            "",
            f"<b>{number}.</b> {html.escape(variant.name)}",
            f"📍 {html.escape(warehouses)}",
        ])
    return "\n".join(lines)


def price_variants_keyboard(callback_id: int, page: int, total: int) -> InlineKeyboardMarkup:
    page_count = max(1, (total + PRICE_VARIANTS_PER_PAGE - 1) // PRICE_VARIANTS_PER_PAGE)
    navigation = []
    if page > 0:
        navigation.append(
            InlineKeyboardButton(text="⬅️ Назад", callback_data=f"price:v:{callback_id}:{page - 1}")
        )
    if page + 1 < page_count:
        navigation.append(
            InlineKeyboardButton(text="Далее ➡️", callback_data=f"price:v:{callback_id}:{page + 1}")
        )
    rows = [navigation] if navigation else []
    rows.extend([
        [InlineKeyboardButton(text="💰 Вернуться к ценам", callback_data=f"price:g:{callback_id}")],
        [InlineKeyboardButton(text="🔎 Новый поиск", callback_data="main:prices")],
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.callback_query(F.data == "main:prices")
async def open_price_search(callback: CallbackQuery, state: FSMContext) -> None:
    if not prices_db.import_statuses():
        await callback.answer("Прайсы пока не загружены.", show_alert=True)
        return
    await state.set_state(AppState.price_search)
    await callback.message.edit_text(
        "💰 <b>Цены и наличие</b>\n\n"
        "Введите название товарной группы. Можно использовать бренд, модель, категорию "
        "или их часть — точное совпадение не требуется.\n\n"
        "Например: <code>OGGO VLIQ</code> или <code>Dojo 12000</code>",
        reply_markup=back_main(),
    )
    await callback.answer()


@router.message(AppState.price_search, F.text)
async def search_prices(message: Message) -> None:
    query = message.text.strip()
    if len(query) > 200:
        await message.answer("⚠️ Запрос слишком длинный. Укажите только товарную группу.", reply_markup=back_main())
        return
    groups = prices_db.search_groups(query)
    if not groups:
        await message.answer(
            "🤷 <b>Товарная группа не найдена</b>\n\n"
            "Попробуйте сократить запрос, проверить название бренда или указать модель.",
            reply_markup=back_main(),
        )
        return
    rows = [
        [
            InlineKeyboardButton(
                text=price_group_label(group.display_name, group.category_name),
                callback_data=f"price:g:{group.callback_id}",
            )
        ]
        for group in groups
    ]
    rows.extend([
        [InlineKeyboardButton(text="🔎 Новый поиск", callback_data="main:prices")],
        [InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")],
    ])
    await message.answer(
        f"🔎 <b>Подходящих групп: {len(groups)}</b>\n\nВыберите нужную:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )


@router.message(AppState.price_search)
async def price_non_text(message: Message) -> None:
    await message.answer("Отправьте название товара обычным текстом.", reply_markup=back_main())


@router.callback_query(F.data.startswith("price:g:"))
async def show_price_group(callback: CallbackQuery) -> None:
    try:
        callback_id = int(callback.data.rsplit(":", 1)[1])
    except ValueError:
        await callback.answer("Некорректный запрос.", show_alert=True)
        return
    details = prices_db.group_details(callback_id)
    if not details:
        await callback.answer("Прайс обновился. Выполните поиск ещё раз.", show_alert=True)
        return
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🎨 Вкусы и цвета", callback_data=f"price:v:{callback_id}:0")],
        [InlineKeyboardButton(text="🔎 Новый поиск", callback_data="main:prices")],
        [InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")],
    ])
    text = format_price_group(details)
    if len(text) <= 4000:
        await callback.message.edit_text(text, reply_markup=keyboard)
    else:
        await callback.message.edit_text(
            f"💰 <b>{html.escape(details.summary.display_name)}</b>\n\n"
            "В группе много ценовых уровней — отправляю подробный расчёт отдельным сообщением.",
            reply_markup=keyboard,
        )
        await callback.message.answer(text)
    await callback.answer()


@router.callback_query(F.data.startswith("price:v:"))
async def show_price_variants(callback: CallbackQuery) -> None:
    try:
        _, _, raw_id, raw_page = callback.data.split(":", 3)
        callback_id, page = int(raw_id), int(raw_page)
    except ValueError:
        await callback.answer("Некорректный запрос.", show_alert=True)
        return
    result = prices_db.group_variants(callback_id)
    if not result:
        await callback.answer("Прайс обновился. Выполните поиск ещё раз.", show_alert=True)
        return
    summary, variants = result
    await callback.message.edit_text(
        format_price_variants(summary, variants, page),
        reply_markup=price_variants_keyboard(callback_id, page, len(variants)),
    )
    await callback.answer()


def products_keyboard(admin: bool = False) -> InlineKeyboardMarkup:
    products = materials_db.list_products(visible_only=not admin)
    rows = []
    for product in products:
        marker = "🟢" if product.is_visible else "⚪️"
        text = f"{marker} {product.name}" if admin else product.name
        prefix = "adm:p" if admin else "db:p"
        rows.append([InlineKeyboardButton(text=text, callback_data=f"{prefix}:{product.id}")])
    if admin:
        rows.append([InlineKeyboardButton(text="➕ Добавить товар", callback_data="adm:add_product")])
        rows.append([InlineKeyboardButton(text="💰 Обновить прайсы", callback_data="adm:prices")])
        rows.append([InlineKeyboardButton(text="📊 Обновить Excel", callback_data="adm:excel")])
        rows.append([InlineKeyboardButton(text="💾 Скачать резервную копию", callback_data="adm:backup")])
    rows.append([InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.callback_query(F.data == "main:database")
async def open_database(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    products = materials_db.list_products(visible_only=True)
    text = (
        "🗃 <b>База данных</b>\n\nВыберите продукцию, чтобы посмотреть доступные материалы:"
        if products
        else "🗃 <b>База данных</b>\n\nМатериалы пока не опубликованы. Загляните сюда позднее."
    )
    await callback.message.edit_text(text, reply_markup=products_keyboard())
    await callback.answer()


@router.callback_query(F.data.startswith("db:p:"))
async def open_product(callback: CallbackQuery) -> None:
    product = materials_db.get_product(int(callback.data.rsplit(":", 1)[1]))
    if not product or not product.is_visible:
        await callback.answer("Товар недоступен.", show_alert=True)
        return
    sections = materials_db.list_sections(product.id)
    rows = [[InlineKeyboardButton(text=s.name, callback_data=f"db:s:{s.id}")] for s in sections]
    rows.append([InlineKeyboardButton(text="⬅️ К товарам", callback_data="main:database")])
    text = (
        f"📦 <b>{html.escape(product.name)}</b>\n\nВыберите нужный раздел:"
        if sections
        else f"📦 <b>{html.escape(product.name)}</b>\n\nВ этом разделе пока нет опубликованных материалов."
    )
    await callback.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=rows))
    await callback.answer()


async def send_material(bot: Bot, chat_id: int, material: Material) -> None:
    if material.kind == "text":
        await bot.send_message(chat_id, material.text or "", parse_mode=None)
    elif material.kind == "photo":
        await bot.send_photo(chat_id, material.file_id, caption=material.caption, parse_mode=None)
    elif material.kind == "document":
        await bot.send_document(chat_id, material.file_id, caption=material.caption, parse_mode=None)


@router.callback_query(F.data.startswith("db:s:"))
async def deliver_section(callback: CallbackQuery, bot: Bot) -> None:
    section = materials_db.get_section(int(callback.data.rsplit(":", 1)[1]))
    if not section:
        await callback.answer("Раздел не найден.", show_alert=True)
        return
    product = materials_db.get_product(section.product_id)
    if not product or not product.is_visible:
        await callback.answer("Материал недоступен.", show_alert=True)
        return
    items = materials_db.list_materials(section.id)
    if not items:
        await callback.answer("В этом разделе пока нет опубликованных материалов.", show_alert=True)
        return
    await callback.answer("Материалы отправляются…")
    await bot.send_message(
        callback.message.chat.id,
        f"📎 <b>{html.escape(product.name)}</b>\nРаздел: <b>{html.escape(section.name)}</b>",
    )
    for item in items:
        try:
            await send_material(bot, callback.message.chat.id, item)
        except Exception:
            logging.exception("Не удалось отправить материал %s", item.id)
            await bot.send_message(
                callback.message.chat.id,
                "⚠️ Один из материалов временно недоступен. Сообщите об этом администратору.",
            )


async def require_admin(callback: CallbackQuery) -> bool:
    if is_admin(callback.from_user.id) and callback.message.chat.type == "private":
        return True
    await callback.answer("Этот раздел доступен только администратору.", show_alert=True)
    return False


async def cleanup_pending_excel(state: FSMContext) -> None:
    data = await state.get_data()
    for key in ("pending_excel", "pending_price"):
        pending = data.get(key)
        if pending:
            Path(pending).unlink(missing_ok=True)


def build_backup_archive(archive_path: Path) -> None:
    with tempfile.TemporaryDirectory() as temp_name:
        temp_dir = Path(temp_name)
        database_copy = temp_dir / "materials.sqlite3"
        excel_copy = temp_dir / "managers.xlsx"
        materials_db.backup_to(database_copy)
        shutil.copy2(active_excel_path, excel_copy)
        prices_database = globals().get("prices_db")
        prices_copy = temp_dir / "prices.sqlite3"
        if prices_database is not None:
            prices_database.backup_to(prices_copy)
        metadata = temp_dir / "README.txt"
        metadata.write_text(
            "Резервная копия Ural Vape Regions Bot\n"
            f"Создана: {datetime.now().astimezone().isoformat(timespec='seconds')}\n"
            f"Записей территорий: {len(catalog.entries)}\n"
            "materials.sqlite3 — товары, разделы и материалы\n"
            "managers.xlsx — действующая таблица территорий\n"
            "prices.sqlite3 — загруженные складские цены и товарные группы\n"
            "price_files/ — последние исходные прайсы складов\n",
            encoding="utf-8",
        )
        with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.write(database_copy, database_copy.name)
            archive.write(excel_copy, excel_copy.name)
            if prices_copy.exists():
                archive.write(prices_copy, prices_copy.name)
            storage = globals().get("price_storage_path")
            if storage and storage.exists():
                for warehouse in WAREHOUSES:
                    source = storage / f"{warehouse}.xlsx"
                    if source.exists():
                        archive.write(source, f"price_files/{source.name}")
            archive.write(metadata, metadata.name)


@router.callback_query(F.data == "adm:backup")
async def download_backup(callback: CallbackQuery) -> None:
    if not await require_admin(callback):
        return
    await callback.answer("Готовлю резервную копию…")
    await callback.message.answer(
        "⏳ <b>Создаю резервную копию</b>\n\n"
        "Это может занять несколько секунд. Не закрывайте чат."
    )
    try:
        with tempfile.TemporaryDirectory() as temp_name:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            archive_path = Path(temp_name) / f"ural-vape-bot-backup_{timestamp}.zip"
            await asyncio.to_thread(build_backup_archive, archive_path)
            await callback.message.answer_document(
                FSInputFile(archive_path),
                caption=(
                    "✅ <b>Резервная копия готова</b>\n\n"
                    "В архиве находятся база материалов, таблица территорий и складские прайсы. "
                    "Храните файл в надёжном месте."
                ),
            )
    except Exception:
        logging.exception("Не удалось создать резервную копию")
        await callback.message.answer(
            "⚠️ Не удалось создать резервную копию. Попробуйте ещё раз или проверьте журнал сервера."
        )


@router.callback_query(F.data == "adm:excel")
async def start_excel_upload(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await cleanup_pending_excel(state)
    await state.set_state(AdminState.excel_upload)
    await callback.message.edit_text(
        "📊 <b>Обновление таблицы территорий</b>\n\n"
        "Отправьте Excel-файл в формате <code>.xlsx</code>. В первой строке должны быть колонки:\n\n"
        "• <b>Местоположение</b>\n"
        "• <b>Территория</b>\n"
        "• <b>Менеджер</b>\n\n"
        "Сначала бот проверит файл и покажет сводку. Действующая таблица не изменится без подтверждения.",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[[InlineKeyboardButton(text="❌ Отмена", callback_data="main:admin")]]
        ),
    )
    await callback.answer()


@router.message(AdminState.excel_upload)
async def receive_excel(message: Message, state: FSMContext, bot: Bot) -> None:
    if not is_admin(message.from_user.id) or message.chat.type != "private":
        return
    if not message.document or not (message.document.file_name or "").lower().endswith(".xlsx"):
        await message.answer(
            "⚠️ Нужен документ в формате <code>.xlsx</code>. Отправьте правильный файл или нажмите /cancel."
        )
        return
    pending_dir = managed_excel_path.parent / "pending"
    pending_dir.mkdir(parents=True, exist_ok=True)
    pending_path = pending_dir / f"excel_{message.document.file_unique_id}.xlsx"
    await message.answer("⏳ Файл получен. Проверяю структуру и данные…")
    try:
        await bot.download(message.document.file_id, destination=pending_path)
        checked_catalog, details = await asyncio.to_thread(validate_excel, pending_path)
    except Exception as error:
        pending_path.unlink(missing_ok=True)
        logging.warning("Отклонён Excel от администратора: %s", error)
        await message.answer(
            "❌ <b>Файл не прошёл проверку</b>\n\n"
            f"Причина: {html.escape(str(error))}\n\n"
            "Действующая таблица не изменена."
        )
        return
    await state.set_state(AdminState.excel_confirmation)
    await state.update_data(
        pending_excel=str(pending_path),
        excel_rows=len(checked_catalog.entries),
        excel_sheet=details[0],
        excel_name=message.document.file_name,
    )
    await message.answer(
        "✅ <b>Файл успешно проверен</b>\n\n"
        f"Файл: <code>{html.escape(message.document.file_name)}</code>\n"
        f"Лист: <b>{html.escape(details[0])}</b>\n"
        f"Корректных записей: <b>{len(checked_catalog.entries):,}</b>\n\n"
        "Применить эту таблицу? Текущая версия будет сохранена в резервную копию.",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[
                [InlineKeyboardButton(text="✅ Применить", callback_data="adm:apply_excel")],
                [InlineKeyboardButton(text="❌ Отмена", callback_data="adm:cancel_excel")],
            ]
        ),
    )


def apply_pending_excel(pending_path: Path) -> tuple[Catalog, Path]:
    checked_catalog, _ = validate_excel(pending_path)
    backup_dir = managed_excel_path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_path = backup_dir / f"managers_{timestamp}.xlsx"
    shutil.copy2(active_excel_path, backup_path)
    managed_excel_path.parent.mkdir(parents=True, exist_ok=True)
    os.replace(pending_path, managed_excel_path)
    checked_catalog.path = managed_excel_path
    return checked_catalog, backup_path


@router.callback_query(F.data == "adm:apply_excel")
async def apply_excel(callback: CallbackQuery, state: FSMContext) -> None:
    global catalog, active_excel_path
    if not await require_admin(callback):
        return
    data = await state.get_data()
    pending_path = Path(data.get("pending_excel", ""))
    if not pending_path.is_file():
        await state.clear()
        await callback.answer("Файл проверки не найден. Загрузите его ещё раз.", show_alert=True)
        return
    await callback.answer("Применяю таблицу…")
    try:
        new_catalog, backup_path = await asyncio.to_thread(apply_pending_excel, pending_path)
        catalog = new_catalog
        active_excel_path = managed_excel_path
    except Exception:
        logging.exception("Не удалось применить новый Excel")
        await callback.message.edit_text(
            "❌ Не удалось применить таблицу. Действующая версия сохранена без изменений.",
            reply_markup=products_keyboard(admin=True),
        )
        return
    await state.clear()
    await callback.message.edit_text(
        "✅ <b>Таблица обновлена</b>\n\n"
        f"Загружено записей: <b>{len(catalog.entries):,}</b>\n"
        "Новые данные уже используются в поиске — перезапуск бота не требуется.\n\n"
        f"Предыдущая версия сохранена: <code>{html.escape(backup_path.name)}</code>",
        reply_markup=products_keyboard(admin=True),
    )


@router.callback_query(F.data == "adm:cancel_excel")
async def cancel_excel(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await cleanup_pending_excel(state)
    await state.clear()
    await callback.message.edit_text(
        "Обновление таблицы отменено. Действующие данные не изменились.",
        reply_markup=products_keyboard(admin=True),
    )
    await callback.answer()


def price_admin_keyboard() -> InlineKeyboardMarkup:
    statuses = prices_db.import_statuses()
    rows = []
    for warehouse in ("center", "west", "ural"):
        marker = "✅" if warehouse in statuses else "➕"
        rows.append([
            InlineKeyboardButton(
                text=f"{marker} {WAREHOUSES[warehouse]}",
                callback_data=f"adm:price_wh:{warehouse}",
            )
        ])
    rows.append([InlineKeyboardButton(text="⬅️ К управлению", callback_data="main:admin")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def price_status_text() -> str:
    statuses = prices_db.import_statuses()
    lines = [
        "💰 <b>Управление прайсами</b>",
        "",
        "Выберите склад, для которого хотите загрузить свежий прайс.",
        "Данные остальных складов не изменятся.",
        "",
    ]
    for warehouse in ("center", "west", "ural"):
        status = statuses.get(warehouse)
        if status:
            date = (status["price_date"] or status["updated_at"] or "").split("T", 1)[0]
            lines.append(
                f"✅ <b>{WAREHOUSES[warehouse]}</b> — {status['item_count']} позиций, "
                f"прайс от {html.escape(date)}"
            )
        else:
            lines.append(f"➕ <b>{WAREHOUSES[warehouse]}</b> — прайс не загружен")
    return "\n".join(lines)


@router.callback_query(F.data == "adm:prices")
async def admin_prices(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await cleanup_pending_excel(state)
    await state.clear()
    await callback.message.edit_text(price_status_text(), reply_markup=price_admin_keyboard())
    await callback.answer()


@router.callback_query(F.data.startswith("adm:price_wh:"))
async def start_price_upload(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    warehouse = callback.data.rsplit(":", 1)[1]
    if warehouse not in WAREHOUSES:
        await callback.answer("Неизвестный склад.", show_alert=True)
        return
    await cleanup_pending_excel(state)
    await state.set_state(AdminState.price_upload)
    await state.update_data(price_warehouse=warehouse)
    await callback.message.edit_text(
        f"💰 <b>Обновление прайса</b>\n"
        f"Склад: <b>{WAREHOUSES[warehouse]}</b>\n\n"
        "Отправьте свежий прайс в формате <code>.xlsx</code>. Бот проверит структуру, "
        "товарные группы и цены, после чего покажет сводку перед применением.",
        reply_markup=InlineKeyboardMarkup(
            inline_keyboard=[[InlineKeyboardButton(text="❌ Отмена", callback_data="adm:prices")]]
        ),
    )
    await callback.answer()


@router.message(AdminState.price_upload)
async def receive_price_file(message: Message, state: FSMContext, bot: Bot) -> None:
    if not is_admin(message.from_user.id) or message.chat.type != "private":
        return
    data = await state.get_data()
    warehouse = data.get("price_warehouse")
    if warehouse not in WAREHOUSES:
        await state.clear()
        await message.answer("Склад не выбран. Начните загрузку заново.")
        return
    if not message.document or not (message.document.file_name or "").lower().endswith(".xlsx"):
        await message.answer("⚠️ Отправьте прайс как документ в формате <code>.xlsx</code>.")
        return
    pending_dir = price_storage_path / "pending"
    pending_dir.mkdir(parents=True, exist_ok=True)
    pending_path = pending_dir / f"{warehouse}_{message.document.file_unique_id}.xlsx"
    await message.answer("⏳ Прайс получен. Проверяю группы, цены и ассортимент…")
    try:
        await bot.download(message.document.file_id, destination=pending_path)
        parsed = await asyncio.to_thread(parse_price_file, pending_path)
    except Exception as error:
        pending_path.unlink(missing_ok=True)
        logging.warning("Отклонён прайс склада %s: %s", warehouse, error)
        await message.answer(
            "❌ <b>Прайс не прошёл проверку</b>\n\n"
            f"Причина: {html.escape(str(error))}\n\n"
            "Действующие цены не изменены."
        )
        return
    await state.set_state(AdminState.price_confirmation)
    await state.update_data(
        pending_price=str(pending_path),
        price_warehouse=warehouse,
        price_file_name=message.document.file_name,
    )
    price_date = parsed.price_date.split("T", 1)[0] if parsed.price_date else "не указана"
    await message.answer(
        "✅ <b>Прайс успешно проверен</b>\n\n"
        f"Склад: <b>{WAREHOUSES[warehouse]}</b>\n"
        f"Дата прайса: <b>{html.escape(price_date)}</b>\n"
        f"Товарных групп: <b>{len(parsed.groups)}</b>\n"
        f"Товарных позиций: <b>{parsed.item_count}</b>\n"
        f"Позиций с пометкой «АКЦИЯ»: <b>{parsed.action_count}</b>\n\n"
        "Применить этот прайс? Предыдущая версия выбранного склада будет сохранена.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Применить", callback_data="adm:apply_price")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="adm:cancel_price")],
        ]),
    )


def apply_pending_price(
    pending_path: Path, warehouse: str, file_name: str
) -> tuple[int, int, Path | None]:
    parsed = parse_price_file(pending_path)
    backup_dir = price_storage_path / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    database_backup = backup_dir / f"prices_before_{warehouse}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.sqlite3"
    prices_db.backup_to(database_backup)
    source_backup = save_price_source(pending_path, price_storage_path, warehouse)
    prices_db.replace_warehouse(warehouse, parsed, file_name)
    pending_path.unlink(missing_ok=True)
    return len(parsed.groups), parsed.item_count, source_backup


@router.callback_query(F.data == "adm:apply_price")
async def apply_price(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    data = await state.get_data()
    warehouse = data.get("price_warehouse")
    pending_path = Path(data.get("pending_price", ""))
    if warehouse not in WAREHOUSES or not pending_path.is_file():
        await state.clear()
        await callback.answer("Файл проверки не найден. Загрузите прайс ещё раз.", show_alert=True)
        return
    await callback.answer("Применяю прайс…")
    try:
        group_count, item_count, _ = await asyncio.to_thread(
            apply_pending_price, pending_path, warehouse, data.get("price_file_name", pending_path.name)
        )
    except Exception:
        logging.exception("Не удалось применить прайс склада %s", warehouse)
        await callback.message.edit_text(
            "❌ Не удалось применить прайс. Предыдущие данные сохранены.",
            reply_markup=price_admin_keyboard(),
        )
        return
    await state.clear()
    await callback.message.edit_text(
        "✅ <b>Прайс обновлён</b>\n\n"
        f"Склад: <b>{WAREHOUSES[warehouse]}</b>\n"
        f"Товарных групп: <b>{group_count}</b>\n"
        f"Товарных позиций: <b>{item_count}</b>\n\n"
        "Новые цены и наличие уже доступны менеджерам. Перезапуск бота не требуется.",
        reply_markup=price_admin_keyboard(),
    )


@router.callback_query(F.data == "adm:cancel_price")
async def cancel_price(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await cleanup_pending_excel(state)
    await state.clear()
    await callback.message.edit_text(
        "Загрузка прайса отменена. Действующие цены не изменились.",
        reply_markup=price_admin_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data == "main:admin")
async def admin_menu(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await state.clear()
    await callback.message.edit_text(
        "⚙️ <b>Управление базой</b>\n\n"
        "Создавайте товары, добавляйте разделы и наполняйте их материалами. "
        "Скрытые товары видны здесь, но недоступны пользователям.",
        reply_markup=products_keyboard(admin=True),
    )
    await callback.answer()


@router.callback_query(F.data == "adm:add_product")
async def admin_add_product(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await state.set_state(AdminState.product_name)
    await callback.message.edit_text(
        "➕ <b>Новый товар</b>\n\nВведите название так, как его должны видеть пользователи:",
        reply_markup=back_main(),
    )
    await callback.answer()


@router.message(AdminState.product_name, F.text)
async def save_product(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id):
        return
    try:
        product = materials_db.add_product(message.text)
    except sqlite3.IntegrityError:
        await message.answer("⚠️ Товар с таким названием уже существует. Введите другое название.")
        return
    await state.clear()
    await message.answer(
        f"✅ Товар <b>{html.escape(product.name)}</b> создан.\n\nТеперь добавьте в него первый раздел.",
        reply_markup=admin_product_keyboard(product.id),
    )


def admin_product_keyboard(product_id: int) -> InlineKeyboardMarkup:
    product = materials_db.get_product(product_id)
    sections = materials_db.list_sections(product_id)
    rows = [[InlineKeyboardButton(text=f"📁 {s.name}", callback_data=f"adm:s:{s.id}")] for s in sections]
    rows.extend([
        [InlineKeyboardButton(text="➕ Добавить раздел", callback_data=f"adm:add_section:{product_id}")],
        [InlineKeyboardButton(text="✏️ Переименовать", callback_data=f"adm:rename_product:{product_id}")],
        [InlineKeyboardButton(text="🙈 Скрыть" if product and product.is_visible else "👁 Показать", callback_data=f"adm:toggle_product:{product_id}")],
        [InlineKeyboardButton(text="🗑 Удалить товар", callback_data=f"adm:confirm_product:{product_id}")],
        [InlineKeyboardButton(text="⬅️ К управлению", callback_data="main:admin")],
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.callback_query(F.data.startswith("adm:p:"))
async def admin_product(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    product_id = int(callback.data.rsplit(":", 1)[1])
    product = materials_db.get_product(product_id)
    if not product:
        await callback.answer("Товар не найден.", show_alert=True); return
    status = "доступен пользователям" if product.is_visible else "скрыт"
    await callback.message.edit_text(
        f"📦 <b>{html.escape(product.name)}</b>\n\nСтатус: <b>{status}</b>\n"
        f"Разделов: <b>{len(materials_db.list_sections(product_id))}</b>",
        reply_markup=admin_product_keyboard(product_id),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("adm:add_section:"))
async def admin_add_section(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    product_id = int(callback.data.rsplit(":", 1)[1])
    await state.set_state(AdminState.section_name)
    await state.update_data(product_id=product_id)
    await callback.message.edit_text(
        "➕ <b>Новый раздел</b>\n\n"
        "Введите название, например: <code>Декларации</code>, <code>Мокапы</code> или <code>КП</code>.",
        reply_markup=back_main(),
    )
    await callback.answer()


@router.message(AdminState.section_name, F.text)
async def save_section(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    product_id = (await state.get_data())["product_id"]
    try:
        section = materials_db.add_section(product_id, message.text)
    except sqlite3.IntegrityError:
        await message.answer("⚠️ Раздел с таким названием уже существует."); return
    await state.clear()
    await message.answer(
        f"✅ Раздел <b>{html.escape(section.name)}</b> создан.\n\nТеперь можно добавить материалы.",
        reply_markup=admin_section_keyboard(section.id),
    )


def material_title(material: Material, number: int) -> str:
    if material.kind == "text":
        preview = (material.text or "").replace("\n", " ")[:30]
        return f"🗑 {number}. Текст: {preview}"
    if material.kind == "photo":
        return f"🗑 {number}. Изображение"
    return f"🗑 {number}. {material.file_name or 'Файл'}"


def admin_section_keyboard(section_id: int) -> InlineKeyboardMarkup:
    section = materials_db.get_section(section_id)
    items = materials_db.list_materials(section_id)
    rows = [[InlineKeyboardButton(text=material_title(m, i), callback_data=f"adm:confirm_material:{m.id}")] for i, m in enumerate(items, 1)]
    rows.extend([
        [InlineKeyboardButton(text="➕ Добавить материалы", callback_data=f"adm:add_material:{section_id}")],
        [InlineKeyboardButton(text="👁 Просмотреть", callback_data=f"adm:preview:{section_id}")],
        [InlineKeyboardButton(text="✏️ Переименовать", callback_data=f"adm:rename_section:{section_id}")],
        [InlineKeyboardButton(text="🗑 Удалить раздел", callback_data=f"adm:confirm_section:{section_id}")],
        [InlineKeyboardButton(text="⬅️ К товару", callback_data=f"adm:p:{section.product_id}")],
    ])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.callback_query(F.data.startswith("adm:s:"))
async def admin_section(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1])
    section = materials_db.get_section(section_id)
    if not section:
        await callback.answer("Раздел не найден.", show_alert=True); return
    count = len(materials_db.list_materials(section_id))
    await callback.message.edit_text(
        f"📁 <b>{html.escape(section.name)}</b>\n\nМатериалов внутри: <b>{count}</b>\n"
        "Нажмите на материал с символом 🗑, чтобы удалить его.",
        reply_markup=admin_section_keyboard(section_id),
    )
    await callback.answer()


def upload_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="✅ Завершить", callback_data="adm:finish_upload")],
        [InlineKeyboardButton(text="❌ Отменить", callback_data="main:admin")],
    ])


@router.callback_query(F.data.startswith("adm:add_material:"))
async def start_upload(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1])
    await state.set_state(AdminState.material_upload)
    await state.update_data(section_id=section_id)
    await callback.message.edit_text(
        "📎 <b>Добавление материалов</b>\n\n"
        "Отправляйте текст, изображения или документы по одному сообщению. "
        "Пользователь получит их в том же порядке.\n\n"
        "Когда всё будет добавлено, нажмите <b>«Завершить»</b>.",
        reply_markup=upload_keyboard(),
    )
    await callback.answer()


@router.message(AdminState.material_upload)
async def save_material(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    section_id = (await state.get_data())["section_id"]
    if message.text:
        materials_db.add_material(section_id, "text", text=message.text)
        label = "Текст"
    elif message.photo:
        materials_db.add_material(section_id, "photo", file_id=message.photo[-1].file_id, caption=message.caption)
        label = "Изображение"
    elif message.document:
        materials_db.add_material(
            section_id, "document", file_id=message.document.file_id,
            caption=message.caption, file_name=message.document.file_name,
        )
        label = "Файл"
    else:
        await message.answer(
            "⚠️ Этот формат пока не поддерживается. Отправьте текст, изображение или документ.",
            reply_markup=upload_keyboard(),
        ); return
    await message.answer(f"✅ {label} добавлен. Можно отправить следующий материал.", reply_markup=upload_keyboard())


@router.callback_query(F.data == "adm:finish_upload")
async def finish_upload(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    section_id = (await state.get_data()).get("section_id")
    await state.clear()
    if not section_id or not materials_db.get_section(section_id):
        await callback.message.edit_text("Раздел не найден.", reply_markup=products_keyboard(admin=True))
    else:
        await callback.message.edit_text(
            "✅ <b>Материалы сохранены</b>\n\nРаздел готов к использованию.",
            reply_markup=admin_section_keyboard(section_id),
        )
    await callback.answer()


@router.callback_query(F.data.startswith("adm:preview:"))
async def preview_section(callback: CallbackQuery, bot: Bot) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1])
    items = materials_db.list_materials(section_id)
    if not items:
        await callback.answer("Материалов пока нет.", show_alert=True); return
    await callback.answer("Отправляю предпросмотр…")
    for item in items:
        await send_material(bot, callback.message.chat.id, item)


@router.callback_query(F.data.startswith("adm:rename_product:"))
async def ask_product_rename(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    product_id = int(callback.data.rsplit(":", 1)[1])
    await state.set_state(AdminState.product_rename); await state.update_data(product_id=product_id)
    await callback.message.edit_text("✏️ <b>Переименование товара</b>\n\nВведите новое название:", reply_markup=back_main()); await callback.answer()


@router.message(AdminState.product_rename, F.text)
async def save_product_rename(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    product_id = (await state.get_data())["product_id"]
    try: materials_db.rename_product(product_id, message.text)
    except sqlite3.IntegrityError:
        await message.answer("Такое название уже используется."); return
    await state.clear(); await message.answer("✅ Название товара обновлено.", reply_markup=admin_product_keyboard(product_id))


@router.callback_query(F.data.startswith("adm:rename_section:"))
async def ask_section_rename(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1])
    await state.set_state(AdminState.section_rename); await state.update_data(section_id=section_id)
    await callback.message.edit_text("✏️ <b>Переименование раздела</b>\n\nВведите новое название:", reply_markup=back_main()); await callback.answer()


@router.message(AdminState.section_rename, F.text)
async def save_section_rename(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    section_id = (await state.get_data())["section_id"]
    try: materials_db.rename_section(section_id, message.text)
    except sqlite3.IntegrityError:
        await message.answer("Такое название уже используется."); return
    await state.clear(); await message.answer("✅ Название раздела обновлено.", reply_markup=admin_section_keyboard(section_id))


@router.callback_query(F.data.startswith("adm:toggle_product:"))
async def toggle_product(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    product_id = int(callback.data.rsplit(":", 1)[1]); materials_db.toggle_product(product_id)
    await callback.message.edit_reply_markup(reply_markup=admin_product_keyboard(product_id)); await callback.answer("Статус изменён")


def confirm_keyboard(yes_data: str, back_data: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Да, удалить", callback_data=yes_data)],
        [InlineKeyboardButton(text="Отмена", callback_data=back_data)],
    ])


@router.callback_query(F.data.startswith("adm:confirm_product:"))
async def confirm_product(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    product_id = int(callback.data.rsplit(":", 1)[1])
    await callback.message.edit_text(
        "⚠️ <b>Удалить товар?</b>\n\nБудут удалены все его разделы и материалы. Это действие нельзя отменить.",
        reply_markup=confirm_keyboard(f"adm:delete_product:{product_id}", f"adm:p:{product_id}"),
    ); await callback.answer()


@router.callback_query(F.data.startswith("adm:delete_product:"))
async def delete_product(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    materials_db.delete_product(int(callback.data.rsplit(":", 1)[1]))
    await callback.message.edit_text("✅ Товар и все связанные материалы удалены.", reply_markup=products_keyboard(admin=True)); await callback.answer()


@router.callback_query(F.data.startswith("adm:confirm_section:"))
async def confirm_section(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1]); section = materials_db.get_section(section_id)
    await callback.message.edit_text(
        "⚠️ <b>Удалить раздел?</b>\n\nВсе материалы внутри него также будут удалены. Это действие нельзя отменить.",
        reply_markup=confirm_keyboard(f"adm:delete_section:{section_id}", f"adm:s:{section_id}"),
    ); await callback.answer()


@router.callback_query(F.data.startswith("adm:delete_section:"))
async def delete_section(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1]); section = materials_db.get_section(section_id)
    product_id = section.product_id if section else 0; materials_db.delete_section(section_id)
    await callback.message.edit_text("✅ Раздел удалён.", reply_markup=admin_product_keyboard(product_id)); await callback.answer()


@router.callback_query(F.data.startswith("adm:confirm_material:"))
async def confirm_material(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    material_id = int(callback.data.rsplit(":", 1)[1]); material = materials_db.get_material(material_id)
    await callback.message.edit_text(
        "⚠️ <b>Удалить материал?</b>\n\nЭто действие нельзя отменить.",
        reply_markup=confirm_keyboard(f"adm:delete_material:{material_id}", f"adm:s:{material.section_id}"),
    ); await callback.answer()


@router.callback_query(F.data.startswith("adm:delete_material:"))
async def delete_material(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    material_id = int(callback.data.rsplit(":", 1)[1]); material = materials_db.get_material(material_id)
    section_id = material.section_id if material else 0; materials_db.delete_material(material_id)
    await callback.message.edit_text("✅ Материал удалён.", reply_markup=admin_section_keyboard(section_id)); await callback.answer()


@router.message(StateFilter(None))
async def outside_mode(message: Message) -> None:
    await message.answer(
        "Чтобы продолжить, выберите нужный раздел:",
        reply_markup=main_menu(message.from_user.id if message.from_user else None),
    )


async def main() -> None:
    global catalog, materials_db, prices_db, admin_ids
    global active_excel_path, managed_excel_path, price_storage_path
    load_dotenv(BASE_DIR / ".env")
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise RuntimeError("Добавьте BOT_TOKEN в файл .env.")
    excel_path = Path(os.getenv("EXCEL_PATH", "managers.xlsx"))
    managed_excel_path = Path(os.getenv("MANAGED_EXCEL_PATH", "data/managers.xlsx"))
    db_path = Path(os.getenv("MATERIALS_DB", "data/materials.sqlite3"))
    prices_db_path = Path(os.getenv("PRICES_DB", "data/prices.sqlite3"))
    price_storage_path = Path(os.getenv("PRICE_STORAGE", "data/prices"))
    if not excel_path.is_absolute(): excel_path = BASE_DIR / excel_path
    if not managed_excel_path.is_absolute(): managed_excel_path = BASE_DIR / managed_excel_path
    if not db_path.is_absolute(): db_path = BASE_DIR / db_path
    if not prices_db_path.is_absolute(): prices_db_path = BASE_DIR / prices_db_path
    if not price_storage_path.is_absolute(): price_storage_path = BASE_DIR / price_storage_path
    admin_ids = {int(value.strip()) for value in os.getenv("ADMIN_IDS", "5533726476").split(",") if value.strip()}
    active_excel_path = managed_excel_path if managed_excel_path.exists() else excel_path
    catalog = Catalog(active_excel_path)
    materials_db = MaterialsDB(db_path)
    prices_db = PricesDB(prices_db_path)
    bot = Bot(token=token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dispatcher = Dispatcher()
    dispatcher.include_router(router)
    await bot.delete_webhook(drop_pending_updates=False)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    asyncio.run(main())
