"""One-time recovery of users who previously opened the bot in private chat."""

import asyncio
from datetime import datetime
import os
from pathlib import Path

from aiogram import Bot
from aiogram.exceptions import TelegramNetworkError, TelegramRetryAfter, TelegramServerError
from dotenv import load_dotenv
from openpyxl import Workbook

from materials_db import MaterialsDB


BASE_DIR = Path(__file__).resolve().parent


async def inspect_chat(bot: Bot, user_id: int):
    for attempt in range(3):
        try:
            return await bot.get_chat(user_id), ""
        except TelegramRetryAfter as error:
            await asyncio.sleep(error.retry_after + 1)
        except (TelegramNetworkError, TelegramServerError) as error:
            if attempt == 2:
                return None, str(error)
            await asyncio.sleep(2 ** attempt)
        except Exception as error:
            return None, str(error)
    return None, "Не удалось проверить после трёх попыток"


def save_report(path: Path, rows: list[tuple]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Результат"
    sheet.append(["Telegram ID", "Username", "Имя", "Результат", "Ошибка"])
    for row in rows:
        sheet.append(row)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in zip(("A", "B", "C", "D", "E"), (18, 25, 35, 24, 70)):
        sheet.column_dimensions[column].width = width
    workbook.save(path)
    workbook.close()


async def main() -> None:
    load_dotenv(BASE_DIR / ".env")
    token = os.getenv("BOT_TOKEN", "").strip()
    if not token:
        raise RuntimeError("В .env не задан BOT_TOKEN")

    db_path = Path(os.getenv("MATERIALS_DB", "data/materials.sqlite3"))
    if not db_path.is_absolute():
        db_path = BASE_DIR / db_path
    database = MaterialsDB(db_path)
    candidates = database.list_unconfirmed_telegram_users()
    print(f"Неподтверждённых ID: {len(candidates)}")

    report_rows: list[tuple] = []
    confirmed = 0
    bot = Bot(token)
    try:
        for number, (user_id, old_username, old_name) in enumerate(candidates, 1):
            chat, error = await inspect_chat(bot, user_id)
            if chat is not None and chat.type == "private":
                username = chat.username or old_username
                full_name = " ".join(
                    value for value in (chat.first_name, chat.last_name) if value
                ).strip() or old_name
                database.remember_telegram_user(user_id, username, full_name)
                database.mark_telegram_user_activated(user_id)
                result = "Подтверждён личный чат"
                confirmed += 1
            elif chat is not None:
                username = old_username
                full_name = old_name
                result = f"Не личный чат: {chat.type}"
            else:
                username = old_username
                full_name = old_name
                result = "Не подтверждён"
            report_rows.append((user_id, username or "", full_name, result, error))
            if number % 25 == 0 or number == len(candidates):
                print(f"Проверено {number}/{len(candidates)}; подтверждено: {confirmed}")
            await asyncio.sleep(0.06)
    finally:
        await bot.session.close()

    report_dir = BASE_DIR / "data" / "recovery_reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    report_path = report_dir / f"private-users-{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"
    save_report(report_path, report_rows)
    total = len(database.list_activated_telegram_users())
    print(f"Найдено новых личных пользователей: {confirmed}")
    print(f"Всего получателей рассылки: {total}")
    print(f"Отчёт: {report_path}")


if __name__ == "__main__":
    asyncio.run(main())
