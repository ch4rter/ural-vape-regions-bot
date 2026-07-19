import asyncio
import html
import logging
import os
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path

from aiogram import Bot, Dispatcher, F, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.types import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message
from dotenv import load_dotenv
from openpyxl import load_workbook


BASE_DIR = Path(__file__).resolve().parent
router = Router()


def normalize(value: str) -> str:
    value = value.casefold().replace("ё", "е")
    value = re.sub(r"[\s\-–—_,.;:()]+", " ", value)
    return value.strip()


@dataclass(frozen=True)
class Entry:
    name: str
    manager: str


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
        rows = list(sheet.iter_rows(values_only=True))
        workbook.close()

        if not rows:
            raise ValueError("Excel-файл пуст.")

        first = [normalize(str(cell or "")) for cell in rows[0][:2]]
        has_header = bool(first and ("регион" in first[0] or "назван" in first[0]))
        data_rows = rows[1:] if has_header else rows

        entries: list[Entry] = []
        by_name: dict[str, list[int]] = {}
        for row_number, row in enumerate(data_rows, start=2 if has_header else 1):
            name = str(row[0]).strip() if len(row) > 0 and row[0] is not None else ""
            manager = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            if not name and not manager:
                continue
            if not name or not manager:
                logging.warning("Пропущена неполная строка Excel: %s", row_number)
                continue
            index = len(entries)
            entries.append(Entry(name=name, manager=manager))
            by_name.setdefault(normalize(name), []).append(index)

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


def format_result(entries: list[Entry]) -> str:
    unique = list(dict.fromkeys((entry.name, entry.manager) for entry in entries))
    if len(unique) == 1:
        name, manager = unique[0]
        return f"📍 <b>{html.escape(name)}</b>\nМенеджер: <b>{html.escape(manager)}</b>"
    lines = ["Найдено несколько записей:"]
    for name, manager in unique:
        lines.append(f"\n📍 <b>{html.escape(name)}</b>\nМенеджер: <b>{html.escape(manager)}</b>")
    return "\n".join(lines)


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
        await message.answer(format_result(found))
        return

    indexes = catalog.suggestions(query)
    if not indexes:
        await message.answer("Ничего не найдено. Проверьте написание или попробуйте другое название.")
        return

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=catalog.entries[i].name, callback_data=f"place:{i}")]
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
    await callback.message.edit_text(format_result(found))
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
