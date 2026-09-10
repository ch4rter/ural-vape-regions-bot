from decimal import Decimal

from inventory_grouping import InventoryGroupingRule, export_grouping, load_grouping, rule_key, save_grouping
from inventory_inline import (InventoryGroup, InventoryItem, InventorySnapshot,
    inventory_card_text, inventory_cards, search_inventory_lines)


def test_grouping_roundtrip_and_card_counts(tmp_path):
    rule = InventoryGroupingRule("OGGO VLIQ ICE 20 мг", "Жидкости", "OGGO VLIQ", "OGGO VLIQ ICE 20 мг", "Жидкости", "флаконов")
    source = tmp_path / "rules.xlsx"
    from prices_db import ItemSummary
    export_grouping(source, [ItemSummary(1, "Манго", rule.source_group, rule.source_category, {}, "1")], {rule_key(rule.source_group, rule.source_category): rule})
    target = tmp_path / "rules.json"
    assert save_grouping(source, target) == 1
    rules = load_grouping(target)
    items = (
        InventoryItem("1", "Манго", rule.source_group, rule.source_category, (Decimal(10), Decimal(0), Decimal(2))),
        InventoryItem("2", "Ягоды", rule.source_group, rule.source_category, (Decimal(0), Decimal(0), Decimal(0))),
    )
    snapshot = InventorySnapshot(items, (InventoryGroup("g", rule.source_group, rule.source_category, items),))
    card = inventory_cards(snapshot, rules)[0]
    text = inventory_card_text(card, "10.09.2026 12:00")
    assert "10 флаконов" in text
    assert "1/2 вкуса" in text
    assert search_inventory_lines((card,), "OGGO 20 мг")[0].name == "OGGO VLIQ ICE 20 мг"


def test_duplicate_technical_groups_merge_into_one_line():
    first = InventoryItem("1", "Манго", "OGGO X Dojo 10000", "OGGO", (Decimal(1), Decimal(0), Decimal(0)))
    second = InventoryItem("2", "Арбуз", "OGGO x Dojo 10000", "OGGO", (Decimal(0), Decimal(1), Decimal(0)))
    snapshot = InventorySnapshot((first, second), (
        InventoryGroup("1", first.group_name, first.category_name, (first,)),
        InventoryGroup("2", second.group_name, second.category_name, (second,)),
    ))
    rules = {
        rule_key(first.group_name, first.category_name): InventoryGroupingRule(first.group_name, "OGGO", "OGGO x DOJO", "OGGO x DOJO 10000", "Одноразки", "шт."),
        rule_key(second.group_name, second.category_name): InventoryGroupingRule(second.group_name, "OGGO", "OGGO x DOJO", "oggo X dojo 10000", "Одноразки", "шт."),
    }
    card = inventory_cards(snapshot, rules)[0]
    assert len(card.lines) == 1
    assert len(card.lines[0].items) == 2


def test_excluded_group_is_saved_without_card(tmp_path):
    from openpyxl import load_workbook
    from prices_db import ItemSummary
    source = tmp_path / "rules.xlsx"
    export_grouping(source, [ItemSummary(1, "Товар", "Служебная", "Прочее", {}, "1")], {})
    workbook = load_workbook(source)
    workbook["Правила"]["G2"] = "Да"
    workbook.save(source)
    target = tmp_path / "rules.json"
    assert save_grouping(source, target) == 1
    rule = next(iter(load_grouping(target).values()))
    assert rule.excluded is True
    assert rule.card_name == ""
