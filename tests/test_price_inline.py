from decimal import Decimal

from price_inline import (
    grouped_card_price_text,
    grouped_price_cards,
    search_grouped_prices,
    group_price_description,
    group_price_text,
    item_price_description,
    item_price_text,
)
from inventory_grouping import InventoryGroupingRule, rule_key
from prices_db import GroupDetails, GroupSummary, ItemSummary, PriceTier


def test_group_price_with_multiple_tiers():
    summary = GroupSummary(1, "x", "Vaporesso расходники", "Картриджи", "", {})
    details = GroupDetails(summary, (
        PriceTier(Decimal("220"), Decimal("230"), 4),
        PriceTier(Decimal("250.50"), Decimal("260.50"), 2),
    ), 6)
    text = group_price_text(details, "08.09.2026")
    assert "Ценовые уровни" in text
    assert "220 ₽" in text
    assert "260,50 ₽" in text
    assert "актуальны на 08.09.2026" in text
    assert group_price_description(details) == "Нал 220–250,50 ₽ · Безнал 230–260,50 ₽"


def test_item_base_prices():
    item = ItemSummary(
        5, "Картридж XROS 0.6", "Vaporesso расходники", "Картриджи",
        {"common": (Decimal("225"), Decimal("235"))}, "001",
    )
    assert item_price_description(item) == "Нал 225 ₽ · Безнал 235 ₽"
    text = item_price_text(item, "08.09.2026")
    assert "Базовая цена товара" in text
    assert "Нал — <b>225 ₽</b>" in text
    assert "Безнал — <b>235 ₽</b>" in text


def test_grouped_prices_use_inventory_rules_and_keep_price_tiers():
    items = [
        ItemSummary(1, "Манго", "Dojo 10000", "OGGO", {"common": (Decimal("500"), Decimal("520"))}, "1", "Одноразки/OGGO/Dojo 10000", "f1"),
        ItemSummary(2, "Арбуз", "Dojo 10000", "OGGO", {"common": (Decimal("550"), Decimal("570"))}, "2", "Одноразки/OGGO/Dojo 10000", "f1"),
    ]
    rule = InventoryGroupingRule("Dojo 10000", "OGGO", "OGGO x DOJO", "OGGO x DOJO 10000", "Одноразки", "шт.", False, items[0].folder_path, "f1")
    cards = grouped_price_cards(items, {rule_key(rule.source_group, rule.source_category, rule.source_path, rule.source_folder_id): rule})
    found_cards, found_lines = search_grouped_prices(cards, "OGGO DOJO 10000")
    assert found_cards and found_lines
    text = grouped_card_price_text(found_cards[0], "10.09.2026")
    assert "OGGO x DOJO 10000" in text
    assert "500 ₽" in text and "570 ₽" in text
    assert "1 вкус" in text
