"""Formatting helpers for cached inline price lookup."""

from __future__ import annotations

import html
from decimal import Decimal

from prices_db import GroupDetails, ItemSummary, PRICE_SOURCE


def money(value: Decimal) -> str:
    rendered = f"{value:,.2f}".replace(",", " ").replace(".00", "")
    return rendered.replace(".", ",")


def _range(values: list[Decimal]) -> str:
    if not values:
        return "—"
    minimum, maximum = min(values), max(values)
    if minimum == maximum:
        return f"{money(minimum)} ₽"
    return f"{money(minimum)}–{money(maximum)} ₽"


def group_price_description(details: GroupDetails) -> str:
    cash = [tier.cash for tier in details.tiers if tier.cash > 0]
    cashless = [tier.cashless for tier in details.tiers if tier.cashless > 0]
    return f"Нал {_range(cash)} · Безнал {_range(cashless)}"


def group_price_text(details: GroupDetails, updated_at: str) -> str:
    lines = [
        "💰 <b>Базовые цены товарной группы</b>",
        "",
        f"<b>{html.escape(details.summary.display_name)}</b>",
        f"<i>{html.escape(details.summary.category_name)}</i>",
        "",
    ]
    if len(details.tiers) == 1:
        tier = details.tiers[0]
        lines.extend([
            f"• Нал — <b>{money(tier.cash)} ₽</b>",
            f"• Безнал — <b>{money(tier.cashless)} ₽</b>",
        ])
    else:
        lines.append("<b>Ценовые уровни:</b>")
        visible_tiers = details.tiers[:12]
        for index, tier in enumerate(visible_tiers, 1):
            lines.append(
                f"{index}. Нал — <b>{money(tier.cash)} ₽</b> · "
                f"безнал — <b>{money(tier.cashless)} ₽</b>"
            )
        if len(details.tiers) > len(visible_tiers):
            lines.append(f"…ещё {len(details.tiers) - len(visible_tiers)} ценовых уровней")
    lines.extend(("", f"<blockquote>🕒 Цены актуальны на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)


def item_prices(item: ItemSummary) -> tuple[Decimal, Decimal] | None:
    return item.warehouse_prices.get(PRICE_SOURCE)


def item_price_description(item: ItemSummary) -> str:
    prices = item_prices(item)
    if not prices:
        return "Базовая цена не указана"
    return f"Нал {money(prices[0])} ₽ · Безнал {money(prices[1])} ₽"


def item_price_text(item: ItemSummary, updated_at: str) -> str:
    prices = item_prices(item)
    lines = [
        "💰 <b>Базовая цена товара</b>",
        "",
        f"<b>{html.escape(item.name)}</b>",
        f"<i>{html.escape(item.group_name)}</i>",
        "",
    ]
    if prices:
        lines.extend([
            f"• Нал — <b>{money(prices[0])} ₽</b>",
            f"• Безнал — <b>{money(prices[1])} ₽</b>",
        ])
    else:
        lines.append("Цена пока не указана.")
    lines.extend(("", f"<blockquote>🕒 Цены актуальны на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)
