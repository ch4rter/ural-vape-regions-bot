import bot
from bot import (
    HOME_BUTTON_TEXT,
    clean_client_title,
    compact_nav,
    group_welcome_text,
    main_menu,
    persistent_home_keyboard,
)
from materials_db import MaterialsDB


def test_public_main_menu_only_shows_public_sections():
    keyboard = main_menu(None)
    assert [len(row) for row in keyboard.inline_keyboard] == [2]
    assert [button.text for button in keyboard.inline_keyboard[0]] == [
        "🔎 Менеджеры", "🗃 База данных",
    ]


def test_completed_public_profile_gets_base_price_button(tmp_path):
    previous = getattr(bot, "materials_db", None)
    bot.materials_db = MaterialsDB(tmp_path / "materials.sqlite3")
    bot.materials_db.save_lead_profile(
        123456, "client", "Иван", "Владелец", "Тамбов",
        "Vape Shop", "2 точки", "Андрей",
    )
    try:
        labels = [
            button.text
            for row in main_menu(123456).inline_keyboard
            for button in row
        ]
        assert "📄 Получить прайсы" in labels
        assert "💰 Цены" not in labels
    finally:
        if previous is not None:
            bot.materials_db = previous
        else:
            delattr(bot, "materials_db")


def test_compact_navigation_uses_icon_only_buttons():
    row = compact_nav("section:back", forward_data="section:next", search_data="main:prices")
    assert [button.text for button in row] == ["⬅️", "➡️", "🔎", "🏠"]
    assert row[-1].callback_data == "main:menu"


def test_persistent_home_keyboard_and_group_welcome():
    keyboard = persistent_home_keyboard()
    assert keyboard.is_persistent is True
    assert keyboard.keyboard[0][0].text == HOME_BUTTON_TEXT
    welcome = group_welcome_text()
    assert "бот-помощник URAL VAPE" in welcome
    assert "/прайс" in welcome
    assert "мокапы, декларации и промоматериалы" in welcome


def test_client_chat_title_removes_company_name():
    assert clean_client_title("Магазин Табак | URAL VAPE") == "Магазин Табак"
    assert clean_client_title("URAL VAPE — Клиент 24") == "Клиент 24"
