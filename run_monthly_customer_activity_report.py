"""Build a monthly customer activity report and send a test to main admins."""

from __future__ import annotations

import argparse
import asyncio
import os
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from aiogram import Bot
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.types import FSInputFile
from dotenv import load_dotenv

from customer_activity import ActivityDatabase, build_monthly_report
from materials_db import MaterialsDB
from moysklad_price import MoySkladClient


BASE_DIR = Path(__file__).resolve().parent


def resolved_path(setting: str, default: str) -> Path:
    path = Path(os.getenv(setting, default))
    return path if path.is_absolute() else BASE_DIR / path


def previous_month(today: date) -> date:
    current = today.replace(day=1)
    return date(current.year - 1, 12, 1) if current.month == 1 else date(current.year, current.month - 1, 1)


async def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", help="отчётный месяц в формате YYYY-MM")
    parser.add_argument("--full-sync", action="store_true", help="заново прочитать отгрузки с 01.04.2026")
    args = parser.parse_args()
    load_dotenv(BASE_DIR / ".env")
    bot_token = os.getenv("BOT_TOKEN", "").strip()
    moysklad_token = os.getenv("MOYSKLAD_TOKEN", "").strip()
    admin_ids = {
        int(value.strip()) for value in os.getenv("ADMIN_IDS", "5533726476").split(",") if value.strip()
    }
    if not bot_token or not moysklad_token or not admin_ids:
        raise RuntimeError("Проверьте BOT_TOKEN, MOYSKLAD_TOKEN и ADMIN_IDS в .env.")
    try:
        month_start = datetime.strptime(args.month, "%Y-%m").date() if args.month else previous_month(date.today())
    except ValueError as error:
        raise RuntimeError("Месяц должен быть указан в формате YYYY-MM, например 2026-08.") from error

    materials = MaterialsDB(resolved_path("MATERIALS_DB", "data/materials.sqlite3"))
    activity = ActivityDatabase(resolved_path("CUSTOMER_ACTIVITY_DB", "data/customer_activity.sqlite3"))
    if args.full_sync:
        activity.set_meta("shipments_cursor", "")
    channels = {
        str(user.id): user.sales_channel_href
        for user in materials.list_access_users()
        if user.role == "manager" and user.sales_channel_href
    }
    now = datetime.now(ZoneInfo("Europe/Moscow"))
    print(f"Формирую тестовый месячный отчёт за {month_start:%m.%Y}…")
    report = await asyncio.to_thread(
        build_monthly_report,
        activity,
        MoySkladClient(moysklad_token),
        resolved_path("CUSTOMER_ACTIVITY_STORAGE", "data/customer_activity_reports"),
        month_start,
        channels,
        now.replace(tzinfo=None),
        save=False,
    )
    summary = report.summary
    caption = (
        "🧪 <b>Тестовый месячный отчёт по клиентской базе</b>\n\n"
        f"Период: <b>{datetime.fromisoformat(report.period_start):%d.%m.%Y}–"
        f"{datetime.fromisoformat(report.period_end):%d.%m.%Y}</b>\n\n"
        f"👥 Активная база: <b>{summary['active_start']} → {summary['active_end']}</b> "
        f"(<b>{summary['active_change']:+d}</b>)\n"
        f"🆕 Новые: <b>{summary['new']}</b>\n"
        f"🔄 Вернувшиеся: <b>{summary['returned']}</b>\n"
        f"📉 Стали неактивными: <b>{summary['became_inactive']}</b>\n"
        f"🛒 Уникальные покупатели: <b>{summary['unique_buyers']}</b>\n\n"
        "Это ручная проверка. Штатная рассылка второго числа не отключена."
    )
    bot = Bot(bot_token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    try:
        for admin_id in sorted(admin_ids):
            await bot.send_document(admin_id, FSInputFile(report.common_path), caption=caption)
            print(f"Тестовый отчёт отправлен администратору {admin_id}.")
    finally:
        await bot.session.close()


if __name__ == "__main__":
    asyncio.run(main())
