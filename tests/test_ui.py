from bot import clean_client_title, compact_nav, main_menu


def test_main_menu_uses_two_column_grid():
    keyboard = main_menu(None)
    assert [len(row) for row in keyboard.inline_keyboard] == [2, 2, 2]
    assert keyboard.inline_keyboard[0][0].text == "🔎 Менеджеры"
    assert keyboard.inline_keyboard[0][1].text == "💰 Цены"


def test_compact_navigation_uses_icon_only_buttons():
    row = compact_nav("section:back", forward_data="section:next", search_data="main:prices")
    assert [button.text for button in row] == ["⬅️", "➡️", "🔎", "🏠"]
    assert row[-1].callback_data == "main:menu"


def test_client_chat_title_removes_company_name():
    assert clean_client_title("Магазин Табак | URAL VAPE") == "Магазин Табак"
    assert clean_client_title("URAL VAPE — Клиент 24") == "Клиент 24"
