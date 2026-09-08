from decimal import Decimal

from price_inline import (
    group_price_description,
    group_price_text,
    item_price_description,
    item_price_text,
)
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
