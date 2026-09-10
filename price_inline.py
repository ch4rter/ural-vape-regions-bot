"""Formatting helpers for cached inline price lookup."""

from __future__ import annotations

import html
import hashlib
from dataclasses import dataclass
from decimal import Decimal

from prices_db import GroupDetails, ItemSummary, PRICE_SOURCE
from inventory_grouping import InventoryGroupingRule, rule_key
from inventory_inline import assortment_count, _score


@dataclass(frozen=True)
class GroupedPriceLine:
    name: str
    card_name: str
    kind: str
    items: tuple[ItemSummary, ...]

    @property
    def key(self):
        return hashlib.sha256(f"{self.card_name}|{self.name}".encode()).hexdigest()[:16]


@dataclass(frozen=True)
class GroupedPriceCard:
    name: str
    lines: tuple[GroupedPriceLine, ...]

    @property
    def key(self):
        return hashlib.sha256(self.name.encode()).hexdigest()[:16]


def grouped_price_cards(items: list[ItemSummary], rules: dict[str, InventoryGroupingRule]) -> tuple[GroupedPriceCard, ...]:
    cards = {}; metadata = {}
    for item in items:
        rule = rules.get(rule_key(item.group_name, item.category_name, item.folder_path, item.folder_id))
        if rule is None: rule = rules.get(rule_key(item.group_name, item.category_name))
        if not rule or rule.excluded or not item_prices(item): continue
        line_key = rule.line_name.casefold().split()
        line_key = " ".join(line_key)
        cards.setdefault(rule.card_name, {}).setdefault(line_key, []).append(item)
        metadata.setdefault((rule.card_name, line_key), rule)
    result = []
    for card_name, lines in cards.items():
        grouped = tuple(GroupedPriceLine(metadata[(card_name, key)].line_name, card_name,
            metadata[(card_name, key)].kind, tuple(values)) for key, values in lines.items())
        result.append(GroupedPriceCard(card_name, grouped))
    return tuple(result)


def search_grouped_prices(cards: tuple[GroupedPriceCard, ...], query: str):
    card_results = []; line_results = []
    for card in cards:
        score = _score(query, card.name + " " + " ".join(x.name for x in card.lines))
        if score is not None: card_results.append((score, card))
        for line in card.lines:
            score = _score(query, f"{card.name} {line.name} " + " ".join(x.name for x in line.items))
            if score is not None: line_results.append((score, line))
    return ([x for _, x in sorted(card_results, key=lambda x: x[0], reverse=True)[:8]],
            [x for _, x in sorted(line_results, key=lambda x: x[0], reverse=True)[:12]])


def _line_tiers(line: GroupedPriceLine):
    tiers = {}
    for item in line.items:
        prices = item_prices(item)
        if prices: tiers.setdefault(prices, set()).add(item.code or item.name.casefold())
    return sorted((cash, cashless, len(names)) for (cash, cashless), names in tiers.items())


def grouped_line_price_text(line: GroupedPriceLine, updated_at: str, include_card: bool = True) -> str:
    lines = (["💰 <b>Базовые цены линейки</b>", "", f"<b>{html.escape(line.card_name)}</b>",
              f"<b>{html.escape(line.name)}</b>", ""] if include_card else [f"<b>{html.escape(line.name)}</b>"])
    tiers = _line_tiers(line)
    noun = assortment_count(sum(x[2] for x in tiers), line.kind).split(" ", 1)[1]
    for cash, cashless, count in tiers:
        lines.append(f"• {count} {noun}: нал <b>{money(cash)} ₽</b> · безнал <b>{money(cashless)} ₽</b>")
    if include_card: lines.extend(("", f"<blockquote>🕒 Цены актуальны на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)


def grouped_card_price_text(card: GroupedPriceCard, updated_at: str) -> str:
    lines = ["💰 <b>Базовые цены по линейкам</b>", "", f"<b>{html.escape(card.name)}</b>"]
    for line in card.lines: lines.extend(("", grouped_line_price_text(line, updated_at, False)))
    lines.extend(("", f"<blockquote>🕒 Цены актуальны на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)


def grouped_price_description(lines) -> str:
    tiers = [tier for line in lines for tier in _line_tiers(line)]
    cash = [tier[0] for tier in tiers if tier[0] > 0]
    cashless = [tier[1] for tier in tiers if tier[1] > 0]
    return f"Линеек: {len(lines)} · нал {_range(cash)} · безнал {_range(cashless)}"


def search_price_items(items: list[ItemSummary], query: str, limit: int = 12) -> list[ItemSummary]:
    found = []
    for item in items:
        if not item_prices(item): continue
        score = _score(query, f"{item.name} {item.group_name} {item.category_name}")
        if score is not None: found.append((score, item))
    return [item for _, item in sorted(found, key=lambda value: value[0], reverse=True)[:limit]]


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
