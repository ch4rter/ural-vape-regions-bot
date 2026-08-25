import bot
from bot import (
    HOME_BUTTON_TEXT,
    clean_client_title,
    compact_nav,
    confirm_keyboard,
    group_welcome_text,
    main_menu,
    persistent_home_keyboard,
    rps_keyboard,
    rps_result,
    upload_keyboard,
)
from materials_db import MaterialsDB


def test_public_main_menu_only_shows_public_sections():
    keyboard = main_menu(None)
    assert [len(row) for row in keyboard.inline_keyboard] == [2]
    assert [button.text for button in keyboard.inline_keyboard[0]] == [
        "🔎 Менеджеры", "🗃 База данных",
    ]
    assert all(
        button.callback_data != "adm:split_order"
        for row in keyboard.inline_keyboard
        for button in row
    )


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


def test_employee_with_sales_channel_gets_my_shipments_button(tmp_path):
    previous = getattr(bot, "materials_db", None)
    bot.materials_db = MaterialsDB(tmp_path / "materials.sqlite3")
    user = bot.materials_db.add_access_user("123456789")
    bot.materials_db.set_access_sales_channel(
        user.id,
        "Валера",
        "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/valera",
    )
    try:
        buttons = [
            button
            for row in main_menu(123456789).inline_keyboard
            for button in row
        ]
        own = next(button for button in buttons if button.callback_data == "myship:menu")
        assert own.text == "📈 Мои отгрузки"
        beta = next(button for button in buttons if button.callback_data == "adm:split_order")
        assert beta.text == "🧪 Распределить заказ по складам · БЕТА"
    finally:
        if previous is not None:
            bot.materials_db = previous
        else:
            delattr(bot, "materials_db")


def test_compact_navigation_uses_icon_only_buttons():
    row = compact_nav("section:back", forward_data="section:next", search_data="main:prices")
    assert [button.text for button in row] == ["⬅️", "➡️", "🔎", "🏠"]
    assert row[-1].callback_data == "main:menu"


def test_semantic_button_colors_are_used_for_confirmation_and_danger():
    confirmation = confirm_keyboard("confirm", "back").inline_keyboard[0]
    assert [button.style for button in confirmation] == ["success", "danger"]
    upload = upload_keyboard().inline_keyboard[0]
    assert [button.style for button in upload] == ["success", "danger"]


def test_rock_paper_scissors_rules_and_hidden_choice_buttons():
    assert rps_result("rock", "rock") == 0
    assert rps_result("rock", "scissors") == 1
    assert rps_result("rock", "paper") == 2
    assert rps_result("scissors", "paper") == 1
    assert rps_result("paper", "rock") == 1
    buttons = rps_keyboard("secret").inline_keyboard[0]
    assert [button.text for button in buttons] == ["🪨 Камень", "✂️ Ножницы", "📄 Бумага"]
    assert [button.callback_data for button in buttons] == [
        "game:secret:rock", "game:secret:scissors", "game:secret:paper",
    ]


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
