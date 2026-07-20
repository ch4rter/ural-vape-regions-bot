import asyncio
import html
import logging
import os
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from itertools import chain
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.types import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message
from dotenv import load_dotenv
from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
RESULTS_PER_PAGE = 10
router = Router()


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
            raise FileNotFoundError(
                f"Не найден файл {self.path.name}. Создайте его командой: "
                "python create_template.py"
            )

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
            (i for i, value in enumerate(headers) if "назван" in value or "регион" in value),
            None,
        )
        location_idx = next((i for i, value in enumerate(headers) if "местополож" in value), None)
        has_header = manager_idx is not None and (territory_idx is not None or name_idx is not None)

        if has_header:
            name_idx = territory_idx if territory_idx is not None else name_idx
            data_rows = rows
            start_row = 2
        else:
            name_idx = 0
            manager_idx = 1
            location_idx = None
            data_rows = chain([first_row], rows)
            start_row = 1

        entries: list[Entry] = []
        by_name: dict[str, list[int]] = {}
        for row_number, row in enumerate(data_rows, start=start_row):
            name = str(row[name_idx]).strip() if len(row) > name_idx and row[name_idx] is not None else ""
            manager = (
                str(row[manager_idx]).strip()
                if len(row) > manager_idx and row[manager_idx] is not None
                else ""
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
            entries.append(Entry(name=name, manager=manager, location=location))
            by_name.setdefault(normalize(name), []).append(index)

        workbook.close()
        if not entries:
            raise ValueError("В Excel нет заполненных пар «название — менеджер».")
        self.entries = entries
        self.by_name = by_name
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


catalog: Catalog


def unique_results(entries: list[Entry]) -> list[tuple[str, str, str]]:
    return list(dict.fromkeys((entry.name, entry.location, entry.manager) for entry in entries))


def format_result(entries: list[Entry], page: int = 0) -> str:
    unique = list(dict.fromkeys((entry.name, entry.location, entry.manager) for entry in entries))
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
    shown = unique[start : start + RESULTS_PER_PAGE]
    lines = [f"Найдено вариантов: <b>{len(unique)}</b>"]
    if page_count > 1:
        lines.append(f"Страница <b>{page + 1}</b> из <b>{page_count}</b>")
    for number, (name, location, manager) in enumerate(shown, start=start + 1):
        block = [f"\n<b>{number}.</b> 📍 <b>{html.escape(name)}</b>"]
        if location:
            block.append(f"🗺 Местоположение: <b>{html.escape(location)}</b>")
        block.append(f"👤 Менеджер: <b>{html.escape(manager)}</b>")
        lines.append("\n".join(block))
    return "\n".join(lines)


def suggestion_label(entry: Entry, variant_count: int = 1) -> str:
    if variant_count > 1:
        label = f"{entry.name} — {variant_count} вариантов"
    else:
        label = f"{entry.name} — {entry.location}" if entry.location else entry.name
    return label if len(label) <= 64 else f"{label[:61]}..."


def pagination_keyboard(entry_index: int, entries: list[Entry], page: int) -> InlineKeyboardMarkup | None:
    total = len(unique_results(entries))
    page_count = (total + RESULTS_PER_PAGE - 1) // RESULTS_PER_PAGE
    if page_count <= 1:
        return None
    buttons = []
    if page > 0:
        buttons.append(InlineKeyboardButton(text="⬅️ Назад", callback_data=f"page:{entry_index}:{page - 1}"))
    if page + 1 < page_count:
        buttons.append(InlineKeyboardButton(text="Далее ➡️", callback_data=f"page:{entry_index}:{page + 1}"))
    return InlineKeyboardMarkup(inline_keyboard=[buttons])


@router.message(CommandStart())
@router.message(Command("help"))
async def start(message: Message) -> None:
    await message.answer(
        "Напишите название города, деревни, посёлка, области или другого региона. "
        "Я найду закреплённого менеджера.\n\nНапример: <code>Тамбов</code>"
    )


@router.message(F.text)
async def search(message: Message) -> None:
    query = message.text.strip()
    if query.startswith("/"):
        await message.answer("Неизвестная команда. Просто отправьте название населённого пункта.")
        return
    if len(query) > 200:
        await message.answer("Название слишком длинное. Отправьте только населённый пункт или регион.")
        return

    found = catalog.exact(query)
    if found:
        entry_index = catalog.by_name[normalize(query)][0]
        await message.answer(
            format_result(found), reply_markup=pagination_keyboard(entry_index, found, 0)
        )
        return

    indexes = catalog.suggestions(query)
    if not indexes:
        await message.answer("Ничего не найдено. Проверьте написание или попробуйте другое название.")
        return

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text=suggestion_label(
                        catalog.entries[i], len(catalog.by_name[normalize(catalog.entries[i].name)])
                    ),
                    callback_data=f"place:{i}",
                )
            ]
            for i in indexes
        ]
    )
    await message.answer("Точного совпадения нет. Возможно, вы имели в виду:", reply_markup=keyboard)


@router.callback_query(F.data.startswith("place:"))
async def select_suggestion(callback: CallbackQuery) -> None:
    try:
        index = int(callback.data.split(":", 1)[1])
        selected = catalog.entries[index]
    except (ValueError, IndexError):
        await callback.answer("Список обновился. Выполните поиск ещё раз.", show_alert=True)
        return
    found = catalog.exact(selected.name)
    await callback.message.edit_text(
        format_result(found), reply_markup=pagination_keyboard(index, found, 0)
    )
    await callback.answer()


@router.callback_query(F.data.startswith("page:"))
async def change_page(callback: CallbackQuery) -> None:
    try:
        _, raw_index, raw_page = callback.data.split(":", 2)
        index, page = int(raw_index), int(raw_page)
        selected = catalog.entries[index]
    except (ValueError, IndexError):
        await callback.answer("Список обновился. Выполните поиск ещё раз.", show_alert=True)
        return
    found = catalog.exact(selected.name)
    await callback.message.edit_text(
        format_result(found, page), reply_markup=pagination_keyboard(index, found, page)
    )
    await callback.answer()


@router.message()
async def unsupported(message: Message) -> None:
    await message.answer("Отправьте название обычным текстовым сообщением.")


async def main() -> None:
    global catalog
    load_dotenv(BASE_DIR / ".env")
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise RuntimeError("Добавьте BOT_TOKEN в файл .env (см. .env.example).")

    excel_path = Path(os.getenv("EXCEL_PATH", str(BASE_DIR / "managers.xlsx")))
    if not excel_path.is_absolute():
        excel_path = BASE_DIR / excel_path
    catalog = Catalog(excel_path)

    bot = Bot(token=token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    dispatcher = Dispatcher()
    dispatcher.include_router(router)
    await bot.delete_webhook(drop_pending_updates=False)
    await dispatcher.start_polling(bot)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    asyncio.run(main())
