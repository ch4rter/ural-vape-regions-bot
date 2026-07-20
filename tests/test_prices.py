from decimal import Decimal

from openpyxl import Workbook, load_workbook

from bot import discounted, money, variant_word
from prices_db import PricesDB, clean_group_name, generate_discounted_price, parse_price_file


def make_price(path, suffix="", include_action=True):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["ПРАЙС-ЛИСТ"])
    for _ in range(6):
        sheet.append([])
    sheet.append(["Код", "Наименование", "Ед.изм.", "от 50т.р. нал", "от 50т.р. безнал"])
    sheet.append(["ЭС/3.Жидкости/Производитель OGGO/VLIQ OGGO BALANCE 20"])
    sheet.append(["001", f"VLIQ OGGO BALANCE Манго{suffix}", "шт", 235, 259])
    if include_action:
        sheet.append(["002", "АКЦИЯ VLIQ OGGO BALANCE Арбуз", "шт", 100, 110])
    sheet.append(["ЭС/3.Жидкости/Производитель OGGO/VLIQ OGGO BALANCE 20"])
    sheet.append(["003", f"VLIQ OGGO BALANCE Вишня{suffix}", "шт", 247, 270])
    workbook.save(path)


def test_price_parser_uses_group_rows_and_ignores_actions(tmp_path):
    path = tmp_path / "price.xlsx"
    make_price(path)
    parsed = parse_price_file(path)

    assert len(parsed.groups) == 2
    regular = next(group for group in parsed.groups if group.display_name == "VLIQ OGGO BALANCE 20")
    action = next(group for group in parsed.groups if group.display_name == "VLIQ OGGO BALANCE Арбуз")
    assert regular.category_name == "Жидкости"
    assert len(regular.items) == 2
    assert action.items[0].name == "VLIQ OGGO BALANCE Арбуз"
    assert parsed.action_count == 1


def test_warehouse_replacement_search_and_aggregation(tmp_path):
    center = tmp_path / "center.xlsx"
    west = tmp_path / "west.xlsx"
    make_price(center, include_action=False)
    make_price(west, suffix=" West", include_action=False)
    db = PricesDB(tmp_path / "prices.sqlite3")
    db.replace_warehouse("center", parse_price_file(center), center.name)
    db.replace_warehouse("west", parse_price_file(west), west.name)

    results = db.search_groups("ogo vlq balance")
    assert len(results) == 1
    assert results[0].warehouse_counts == {"center": 2, "west": 2}
    details = db.group_details(results[0].callback_id)
    assert len(details.tiers) == 2
    assert details.unique_variants == 2

    replacement = tmp_path / "replacement.xlsx"
    make_price(replacement, suffix=" New", include_action=False)
    db.replace_warehouse("center", parse_price_file(replacement), replacement.name)
    result = db.search_groups("VLIQ BALANCE")[0]
    assert result.warehouse_counts == {"west": 2, "center": 2}


def test_prices_and_group_name_formatting():
    assert clean_group_name("ЭС/Одноразовые/1. Д Vaporesso Dojo 12000") == "Vaporesso Dojo 12000"
    assert money(discounted(Decimal("235"), 5)) == "223,25"
    assert variant_word(1) == "вариант"
    assert variant_word(2) == "варианта"
    assert variant_word(15) == "вариантов"


def test_discounted_excel_changes_copy_but_not_original(tmp_path):
    source = tmp_path / "base.xlsx"
    destination = tmp_path / "discount.xlsx"
    make_price(source)

    changed = generate_discounted_price(source, destination, 10)
    assert changed == 3

    base = load_workbook(source, data_only=True)
    discounted_book = load_workbook(destination, data_only=True)
    assert base.active["D10"].value == 235
    assert discounted_book.active["D8"].value == "от 50т.р. нал — скидка 10%"
    assert discounted_book.active["E8"].value == "от 50т.р. безнал — скидка 10%"
    assert discounted_book.active["D10"].value == 211.5
    assert discounted_book.active["E10"].value == 233.1
    assert discounted_book.active["D11"].value == 90
    base.close()
    discounted_book.close()
