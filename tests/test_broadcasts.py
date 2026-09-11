import asyncio

from openpyxl import Workbook, load_workbook

from bot import (
    SERVICE_CHAT_ID, build_broadcast_report, build_chats_excel,
    build_bot_users_excel, build_leads_excel, copy_broadcast_source, material_batches,
    parse_audience_excel,
    send_material_batch,
)
import bot
from materials_db import MaterialsDB


class CopyBot:
    def __init__(self):
        self.calls = []

    async def copy_message(self, chat_id, from_chat_id, message_id):
        self.calls.append(("one", chat_id, from_chat_id, [message_id]))

    async def copy_messages(self, chat_id, from_chat_id, message_ids):
        self.calls.append(("many", chat_id, from_chat_id, message_ids))

    async def send_media_group(self, chat_id, media):
        self.calls.append(("album", chat_id, [item.media for item in media]))


def test_audience_excel_parsing(tmp_path):
    assert SERVICE_CHAT_ID == -5565597780
    path = tmp_path / "audience.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Клиент", "Chat ID", "Включить"])
    sheet.append(["Первый", -100123, "Да"])
    sheet.append(["Дубль", -100123, "Да"])
    sheet.append(["Второй", "-100456", "Да"])
    sheet.append(["Исключён", -100789, "Нет"])
    sheet.append(["Ошибка", "abc", "Да"])
    workbook.save(path)

    chat_ids, duplicates, invalid = parse_audience_excel(path)
    assert chat_ids == [-100123, -100456]
    assert duplicates == 1
    assert invalid == 1


def test_chat_export_and_broadcast_report(tmp_path):
    previous = getattr(bot, "materials_db", None)
    bot.materials_db = MaterialsDB(tmp_path / "materials.sqlite3")
    bot.materials_db.upsert_client_chat(-100123, "Клиент", "supergroup", True)
    try:
        chats_path = tmp_path / "chats.xlsx"
        build_chats_excel(chats_path)
        workbook = load_workbook(chats_path, data_only=True)
        assert workbook.active["A2"].value == "Клиент"
        assert workbook.active["B2"].value == -100123
        workbook.close()

        report_path = tmp_path / "report.xlsx"
        build_broadcast_report(report_path, [{
            "title": "Клиент", "username": "@client", "chat_id": 123456,
            "status": "Отправлено", "error": "",
        }])
        report = load_workbook(report_path, data_only=True)
        assert report.active["B2"].value == "@client"
        assert report.active["C2"].value == 123456
        assert report.active["D2"].value == "Отправлено"
        report.close()

        bot.materials_db.remember_telegram_user(987654, "buyer", "Покупатель")
        bot.materials_db.mark_telegram_user_activated(987654)
        users_path = tmp_path / "users.xlsx"
        assert build_bot_users_excel(users_path) == 1
        users = load_workbook(users_path, data_only=True)
        assert users.active["A2"].value == 987654
        assert users.active["B2"].value == "@buyer"
        assert users.active["C2"].value == "Покупатель"
        users.close()

        bot.materials_db.save_lead_profile(
            123456, "client", "Иван", "Владелец", "Тамбов",
            "Vape Shop", "2 точки", "Андрей",
        )
        leads_path = tmp_path / "leads.xlsx"
        assert build_leads_excel(leads_path) == 1
        leads = load_workbook(leads_path, data_only=True)
        assert leads.active["B2"].value == "Иван"
        assert leads.active["I2"].value == 123456
        leads.close()
    finally:
        if previous is not None:
            bot.materials_db = previous


def test_broadcast_album_uses_copy_messages_and_materials_form_one_batch(tmp_path):
    telegram = CopyBot()
    asyncio.run(copy_broadcast_source(telegram, -1001, 42, [13, 11, 12, 12]))
    assert telegram.calls == [("many", -1001, 42, [11, 12, 13])]

    database = MaterialsDB(tmp_path / "materials.sqlite3")
    product = database.add_product("OGGO")
    section = database.add_section(product.id, "Мокапы")
    database.add_material(
        section.id, "photo", file_id="second", media_group_id="album",
        media_group_position=2,
    )
    database.add_material(section.id, "text", text="После альбома")
    database.add_material(
        section.id, "photo", file_id="first", media_group_id="album",
        media_group_position=1,
    )
    batches = material_batches(database.list_materials(section.id))
    assert [[item.file_id or item.text for item in batch] for batch in batches] == [
        ["first", "second"], ["После альбома"]
    ]
    asyncio.run(send_material_batch(telegram, 99, batches[0]))
    assert telegram.calls[-1] == ("album", 99, ["first", "second"])
