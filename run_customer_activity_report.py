"""Manually build the weekly customer activity report and send a test to admins."""

from __future__ import annotations

import asyncio
import os
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

from aiogram import Bot
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.types import FSInputFile
from dotenv import load_dotenv

from customer_activity import ActivityDatabase, build_weekly_report
from materials_db import MaterialsDB
from moysklad_price import MoySkladClient


BASE_DIR = Path(__file__).resolve().parent


def resolved_path(setting: str, default: str) -> Path:
    path = Path(os.getenv(setting, default))
    return path if path.is_absolute() else BASE_DIR / path


async def main() -> None:
    load_dotenv(BASE_DIR / ".env")
    bot_token = os.getenv("BOT_TOKEN", "").strip()
    moysklad_token = os.getenv("MOYSKLAD_TOKEN", "").strip()
    admin_ids = {
        int(value.strip())
        for value in os.getenv("ADMIN_IDS", "5533726476").split(",")
        if value.strip()
    }
    if not bot_token or not moysklad_token or not admin_ids:
        raise RuntimeError("Проверьте BOT_TOKEN, MOYSKLAD_TOKEN и ADMIN_IDS в .env.")

    materials = MaterialsDB(resolved_path("MATERIALS_DB", "data/materials.sqlite3"))
    activity = ActivityDatabase(
        resolved_path("CUSTOMER_ACTIVITY_DB", "data/customer_activity.sqlite3")
    )
    storage = resolved_path("CUSTOMER_ACTIVITY_STORAGE", "data/customer_activity_reports")
    channels = {
        str(user.id): user.sales_channel_href
        for user in materials.list_access_users()
        if user.role == "manager" and user.sales_channel_href
    }
    now = datetime.now(ZoneInfo("Europe/Moscow"))
    monday = now.date() - timedelta(days=now.weekday())
    # A manual check shows the current calendar week. The regular Monday job
    # still reports the fully completed week and saves the official snapshot.
    report_boundary = monday if now.weekday() == 0 else monday + timedelta(days=7)
    print("Синхронизирую отгрузки и формирую отчёт. Первая загрузка может занять несколько минут…")
    report = await asyncio.to_thread(
        build_weekly_report,
        activity,
        MoySkladClient(moysklad_token),
        storage,
        report_boundary,
        channels,
        now.replace(tzinfo=None),
        save=False,
    )
    print(
        f"Получено/обновлено отгрузок: {report.summary.get('synced', 0)}; "
        f"в локальном реестре: {report.summary.get('stored', 0)}."
    )
    def compared(key: str, previous_key: str) -> str:
        current = int(report.summary.get(key, 0))
        previous = int(report.summary.get(previous_key, 0))
        return f"{current} (было {previous}, {current - previous:+d})"
    caption = (
        "🧪 <b>Тестовый еженедельный отчёт</b>\n\n"
        f"Период: <b>{datetime.fromisoformat(report.period_start):%d.%m.%Y}–"
        f"{datetime.fromisoformat(report.period_end):%d.%m.%Y}</b>\n\n"
        f"🆕 Новые кенты: <b>{compared('new', 'previous_new')}</b>\n"
        f"🔄 Вернувшиеся кенты: <b>{compared('returned', 'previous_returned')}</b>\n"
        f"🕒 Кенты-потеряшки: <b>{compared('lost', 'previous_lost')}</b>\n\n"
        "Это ручная проверка. Штатная рассылка в понедельник не отключена."
    )
    bot = Bot(bot_token, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
    try:
        for admin_id in sorted(admin_ids):
            await bot.send_document(
                admin_id, FSInputFile(report.common_path), caption=caption
            )
            print(f"Тестовый отчёт отправлен администратору {admin_id}.")
    finally:
        await bot.session.close()


if __name__ == "__main__":
    asyncio.run(main())
