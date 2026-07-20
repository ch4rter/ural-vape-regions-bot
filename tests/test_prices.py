from decimal import Decimal

from openpyxl import Workbook

from bot import discounted, money, variant_word
from prices_db import PricesDB, clean_group_name, parse_price_file


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

    assert len(parsed.groups) == 1
    assert parsed.groups[0].display_name == "VLIQ OGGO BALANCE 20"
    assert parsed.groups[0].category_name == "Жидкости"
    assert len(parsed.groups[0].items) == 2
    assert parsed.ignored_actions == 1


def test_warehouse_replacement_search_and_aggregation(tmp_path):
    center = tmp_path / "center.xlsx"
    west = tmp_path / "west.xlsx"
    make_price(center)
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
