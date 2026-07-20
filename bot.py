import asyncio
import html
import logging
import os
import re
import sqlite3
from dataclasses import dataclass
from difflib import SequenceMatcher
from itertools import chain
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message
from dotenv import load_dotenv
from openpyxl import load_workbook

from materials_db import Material, MaterialsDB


BASE_DIR = Path(__file__).resolve().parent
RESULTS_PER_PAGE = 10
router = Router()
catalog: "Catalog"
materials_db: MaterialsDB
admin_ids: set[int] = set()


class AppState(StatesGroup):
    region_search = State()


class AdminState(StatesGroup):
    product_name = State()
    product_rename = State()
    section_name = State()
    section_rename = State()
    material_upload = State()


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


def is_admin(user_id: int | None) -> bool:
    return user_id is not None and user_id in admin_ids


def main_menu(user_id: int | None) -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton(text="🔎 Найти менеджера", callback_data="main:region")],
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
    text = "Выберите нужный раздел:"
    if edit:
        await target.edit_text(text, reply_markup=main_menu(user_id))
    else:
        await target.answer(text, reply_markup=main_menu(user_id))


@router.message(CommandStart())
@router.message(Command("menu"))
async def command_menu(message: Message, state: FSMContext) -> None:
    await state.clear()
    await show_main(message, message.from_user.id if message.from_user else None)


@router.message(Command("cancel"))
async def cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Действие отменено.")
    await show_main(message, message.from_user.id if message.from_user else None)


@router.callback_query(F.data == "main:menu")
async def callback_menu(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await show_main(callback.message, callback.from_user.id, edit=True)
    await callback.answer()


@router.callback_query(F.data == "main:region")
async def open_region_search(callback: CallbackQuery, state: FSMContext) -> None:
    await state.set_state(AppState.region_search)
    await callback.message.edit_text(
        "Введите город, деревню, область или другую территорию.\n\nНапример: <code>Питер</code>",
        reply_markup=back_main(),
    )
    await callback.answer()


def unique_results(entries: list[Entry]) -> list[tuple[str, str, str]]:
    return list(dict.fromkeys((entry.name, entry.location, entry.manager) for entry in entries))


def format_result(entries: list[Entry], page: int = 0) -> str:
    unique = unique_results(entries)
    if len(unique) == 1:
        name, location, manager = unique[0]
        lines = [f"📍 <b>{html.escape(name)}</b>"]
        if location:
            lines.append(f"🗺 Местоположение: <b>{html.escape(location)}</b>")
        lines.append(f"👤 Менеджер: <b>{html.escape(manager)}</b>")
        return "\n".join(lines)
    page_count = (len(unique) + RESULTS_PER_PAGE - 1) // RESULTS_PER_PAGE
    page = max(0, min(page, page_count - 1))
    start = page * RESULTS_PER_PAGE
    lines = [f"Найдено вариантов: <b>{len(unique)}</b>"]
    if page_count > 1:
        lines.append(f"Страница <b>{page + 1}</b> из <b>{page_count}</b>")
    for number, (name, location, manager) in enumerate(
        unique[start : start + RESULTS_PER_PAGE], start=start + 1
    ):
        block = [f"\n<b>{number}.</b> 📍 <b>{html.escape(name)}</b>"]
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
        await message.answer("Запрос слишком длинный.", reply_markup=back_main())
        return
    found = catalog.exact(query)
    if found:
        index = catalog.by_name[normalize(query)][0]
        await message.answer(format_result(found), reply_markup=pagination_keyboard(index, found, 0))
        return
    indexes = catalog.suggestions(query)
    if not indexes:
        await message.answer("Ничего не найдено. Проверьте написание.", reply_markup=back_main())
        return
    rows = []
    for index in indexes:
        entry = catalog.entries[index]
        count = len(catalog.by_name[normalize(entry.name)])
        rows.append([InlineKeyboardButton(text=suggestion_label(entry, count), callback_data=f"place:{index}")])
    rows.append([InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")])
    await message.answer(
        "Точного совпадения нет. Возможно, вы имели в виду:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )


@router.message(AppState.region_search)
async def region_non_text(message: Message) -> None:
    await message.answer("Введите название обычным текстом.", reply_markup=back_main())


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
    rows.append([InlineKeyboardButton(text="⬅️ Главное меню", callback_data="main:menu")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.callback_query(F.data == "main:database")
async def open_database(callback: CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    products = materials_db.list_products(visible_only=True)
    text = "Выберите товар:" if products else "База данных пока не заполнена."
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
    text = f"<b>{html.escape(product.name)}</b>\n\nВыберите материал:" if sections else f"<b>{html.escape(product.name)}</b>\n\nМатериалов пока нет."
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
        await callback.answer("В этом разделе пока нет материалов.", show_alert=True)
        return
    await callback.answer("Отправляю материалы…")
    await bot.send_message(callback.message.chat.id, f"<b>{html.escape(product.name)} · {html.escape(section.name)}</b>")
    for item in items:
        try:
            await send_material(bot, callback.message.chat.id, item)
        except Exception:
            logging.exception("Не удалось отправить материал %s", item.id)
            await bot.send_message(callback.message.chat.id, "Не удалось отправить один из материалов.")


async def require_admin(callback: CallbackQuery) -> bool:
    if is_admin(callback.from_user.id) and callback.message.chat.type == "private":
        return True
    await callback.answer("Недостаточно прав.", show_alert=True)
    return False


@router.callback_query(F.data == "main:admin")
async def admin_menu(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await state.clear()
    await callback.message.edit_text("⚙️ <b>Управление базой</b>", reply_markup=products_keyboard(admin=True))
    await callback.answer()


@router.callback_query(F.data == "adm:add_product")
async def admin_add_product(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback):
        return
    await state.set_state(AdminState.product_name)
    await callback.message.edit_text("Введите название нового товара:", reply_markup=back_main())
    await callback.answer()


@router.message(AdminState.product_name, F.text)
async def save_product(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id):
        return
    try:
        product = materials_db.add_product(message.text)
    except sqlite3.IntegrityError:
        await message.answer("Товар с таким названием уже существует.")
        return
    await state.clear()
    await message.answer(f"Товар <b>{html.escape(product.name)}</b> создан.", reply_markup=admin_product_keyboard(product.id))


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
        f"📦 <b>{html.escape(product.name)}</b>\nСтатус: {status}", reply_markup=admin_product_keyboard(product_id)
    )
    await callback.answer()


@router.callback_query(F.data.startswith("adm:add_section:"))
async def admin_add_section(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    product_id = int(callback.data.rsplit(":", 1)[1])
    await state.set_state(AdminState.section_name)
    await state.update_data(product_id=product_id)
    await callback.message.edit_text("Введите название нового раздела:", reply_markup=back_main())
    await callback.answer()


@router.message(AdminState.section_name, F.text)
async def save_section(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    product_id = (await state.get_data())["product_id"]
    try:
        section = materials_db.add_section(product_id, message.text)
    except sqlite3.IntegrityError:
        await message.answer("Раздел с таким названием уже существует."); return
    await state.clear()
    await message.answer(f"Раздел <b>{html.escape(section.name)}</b> создан.", reply_markup=admin_section_keyboard(section.id))


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
        f"📁 <b>{html.escape(section.name)}</b>\nМатериалов: {count}", reply_markup=admin_section_keyboard(section_id)
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
        "Отправляйте текст, изображения или документы по одному сообщению. Они сохранятся в том же порядке.\n\nКогда закончите, нажмите «Завершить».",
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
        await message.answer("Поддерживаются текст, изображения и документы.", reply_markup=upload_keyboard()); return
    await message.answer(f"✅ {label} добавлен.", reply_markup=upload_keyboard())


@router.callback_query(F.data == "adm:finish_upload")
async def finish_upload(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    section_id = (await state.get_data()).get("section_id")
    await state.clear()
    if not section_id or not materials_db.get_section(section_id):
        await callback.message.edit_text("Раздел не найден.", reply_markup=products_keyboard(admin=True))
    else:
        await callback.message.edit_text("Материалы сохранены.", reply_markup=admin_section_keyboard(section_id))
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
    await callback.message.edit_text("Введите новое название товара:", reply_markup=back_main()); await callback.answer()


@router.message(AdminState.product_rename, F.text)
async def save_product_rename(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    product_id = (await state.get_data())["product_id"]
    try: materials_db.rename_product(product_id, message.text)
    except sqlite3.IntegrityError:
        await message.answer("Такое название уже используется."); return
    await state.clear(); await message.answer("Товар переименован.", reply_markup=admin_product_keyboard(product_id))


@router.callback_query(F.data.startswith("adm:rename_section:"))
async def ask_section_rename(callback: CallbackQuery, state: FSMContext) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1])
    await state.set_state(AdminState.section_rename); await state.update_data(section_id=section_id)
    await callback.message.edit_text("Введите новое название раздела:", reply_markup=back_main()); await callback.answer()


@router.message(AdminState.section_rename, F.text)
async def save_section_rename(message: Message, state: FSMContext) -> None:
    if not is_admin(message.from_user.id): return
    section_id = (await state.get_data())["section_id"]
    try: materials_db.rename_section(section_id, message.text)
    except sqlite3.IntegrityError:
        await message.answer("Такое название уже используется."); return
    await state.clear(); await message.answer("Раздел переименован.", reply_markup=admin_section_keyboard(section_id))


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
    await callback.message.edit_text("Удалить товар вместе со всеми разделами и материалами?", reply_markup=confirm_keyboard(f"adm:delete_product:{product_id}", f"adm:p:{product_id}")); await callback.answer()


@router.callback_query(F.data.startswith("adm:delete_product:"))
async def delete_product(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    materials_db.delete_product(int(callback.data.rsplit(":", 1)[1]))
    await callback.message.edit_text("Товар удалён.", reply_markup=products_keyboard(admin=True)); await callback.answer()


@router.callback_query(F.data.startswith("adm:confirm_section:"))
async def confirm_section(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1]); section = materials_db.get_section(section_id)
    await callback.message.edit_text("Удалить раздел и все его материалы?", reply_markup=confirm_keyboard(f"adm:delete_section:{section_id}", f"adm:s:{section_id}")); await callback.answer()


@router.callback_query(F.data.startswith("adm:delete_section:"))
async def delete_section(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    section_id = int(callback.data.rsplit(":", 1)[1]); section = materials_db.get_section(section_id)
    product_id = section.product_id if section else 0; materials_db.delete_section(section_id)
    await callback.message.edit_text("Раздел удалён.", reply_markup=admin_product_keyboard(product_id)); await callback.answer()


@router.callback_query(F.data.startswith("adm:confirm_material:"))
async def confirm_material(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    material_id = int(callback.data.rsplit(":", 1)[1]); material = materials_db.get_material(material_id)
    await callback.message.edit_text("Удалить этот материал?", reply_markup=confirm_keyboard(f"adm:delete_material:{material_id}", f"adm:s:{material.section_id}")); await callback.answer()


@router.callback_query(F.data.startswith("adm:delete_material:"))
async def delete_material(callback: CallbackQuery) -> None:
    if not await require_admin(callback): return
    material_id = int(callback.data.rsplit(":", 1)[1]); material = materials_db.get_material(material_id)
    section_id = material.section_id if material else 0; materials_db.delete_material(material_id)
    await callback.message.edit_text("Материал удалён.", reply_markup=admin_section_keyboard(section_id)); await callback.answer()


@router.message(StateFilter(None))
async def outside_mode(message: Message) -> None:
    await message.answer("Сначала выберите раздел в главном меню.", reply_markup=main_menu(message.from_user.id if message.from_user else None))


async def main() -> None:
    global catalog, materials_db, admin_ids
    load_dotenv(BASE_DIR / ".env")
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise RuntimeError("Добавьте BOT_TOKEN в файл .env.")
    excel_path = Path(os.getenv("EXCEL_PATH", "managers.xlsx"))
    db_path = Path(os.getenv("MATERIALS_DB", "data/materials.sqlite3"))
    if not excel_path.is_absolute(): excel_path = BASE_DIR / excel_path
    if not db_path.is_absolute(): db_path = BASE_DIR / db_path
    admin_ids = {int(value.strip()) for value in os.getenv("ADMIN_IDS", "5533726476").split(",") if value.strip()}
    catalog = Catalog(excel_path)
    materials_db = MaterialsDB(db_path)
    bot = Bot(token=token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dispatcher = Dispatcher()
    dispatcher.include_router(router)
    await bot.delete_webhook(drop_pending_updates=False)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    asyncio.run(main())
