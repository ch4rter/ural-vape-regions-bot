"""Cached, read-only inventory lookup for Telegram inline mode."""

from __future__ import annotations

import hashlib
import html
import json
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from difflib import SequenceMatcher

from moysklad_price import (
    DEFAULT_STORES,
    MoySkladClient,
    _canonical_href,
    _price_type_names,
    _sale_price,
)
from prices_db import ItemSummary, normalize_price_text
from inventory_grouping import InventoryGroupingRule, rule_key


STORE_LABELS = {
    "Мордор": "Москва",
    "Годзибасы": "Санкт-Петербург",
    "Жможики": "Урал",
}


@dataclass(frozen=True)
class InventoryItem:
    key: str
    name: str
    group_name: str
    category_name: str
    quantities: tuple[Decimal, ...]

    @property
    def total(self) -> Decimal:
        return sum(self.quantities, Decimal(0))


@dataclass(frozen=True)
class InventoryGroup:
    key: str
    name: str
    category_name: str
    items: tuple[InventoryItem, ...]

    @property
    def quantities(self) -> tuple[Decimal, ...]:
        return tuple(
            sum((item.quantities[index] for item in self.items), Decimal(0))
            for index in range(len(DEFAULT_STORES))
        )

    @property
    def total(self) -> Decimal:
        return sum(self.quantities, Decimal(0))

    @property
    def variants_by_store(self) -> tuple[int, ...]:
        return tuple(
            sum(1 for item in self.items if item.quantities[index] > 0)
            for index in range(len(DEFAULT_STORES))
        )


@dataclass(frozen=True)
class InventorySnapshot:
    items: tuple[InventoryItem, ...]
    groups: tuple[InventoryGroup, ...]


@dataclass(frozen=True)
class InventoryLine:
    name: str
    kind: str
    unit: str
    items: tuple[InventoryItem, ...]


@dataclass(frozen=True)
class InventoryCard:
    key: str
    name: str
    lines: tuple[InventoryLine, ...]


@dataclass(frozen=True)
class InventoryReference:
    built_on: str
    catalog_signature: str
    store_ids: tuple[str, ...]
    items_by_assortment: dict[str, ItemSummary]


def inventory_catalog_signature(catalog: list[ItemSummary]) -> str:
    source = "\n".join(
        sorted(f"{item.code}|{item.name}|{item.group_name}" for item in catalog)
    )
    return hashlib.sha256(source.encode("utf-8")).hexdigest()


def _store_id(store: dict) -> str:
    return _canonical_href(str(store.get("meta", {}).get("href", ""))).rsplit("/", 1)[-1]


def _assortment_id(item: dict) -> str:
    return _canonical_href(str(item.get("meta", {}).get("href", ""))).rsplit("/", 1)[-1]


def _api_group(assortment: dict) -> tuple[str, str]:
    parts = [value.strip() for value in str(assortment.get("pathName") or "").split("/") if value.strip()]
    if not parts:
        return "Без группы", "Номенклатура"
    return parts[-1], parts[-2] if len(parts) > 1 else "Номенклатура"


def build_inventory_reference(
    client: MoySkladClient,
    catalog: list[ItemSummary],
    *,
    built_on: str | None = None,
    cash_price_type: str = "",
    cashless_price_type: str = "",
) -> InventoryReference:
    stores = client.stores()
    by_name = {str(store.get("name", "")).strip().casefold(): store for store in stores}
    selected = [by_name.get(name.casefold()) for name in DEFAULT_STORES]
    if any(store is None for store in selected):
        missing = [name for name, store in zip(DEFAULT_STORES, selected) if store is None]
        raise RuntimeError(f"В МоемСкладе не найдены склады: {', '.join(missing)}")
    store_ids = tuple(_store_id(store) for store in selected if store is not None)
    catalog_by_code = {item.code.strip().casefold(): item for item in catalog if item.code.strip()}
    catalog_by_name: dict[str, list[ItemSummary]] = {}
    for item in catalog:
        catalog_by_name.setdefault(normalize_price_text(item.name), []).append(item)

    assortment_rows = client.assortment()
    cash_name, cashless_name = _price_type_names(
        assortment_rows, cash_price_type, cashless_price_type
    )
    items_by_assortment = {}
    for assortment in assortment_rows:
        if assortment.get("archived") is True:
            continue
        identifier = _assortment_id(assortment)
        if not identifier:
            continue
        code = str(assortment.get("code") or assortment.get("article") or "").strip().casefold()
        local = catalog_by_code.get(code) if code else None
        if local is None:
            candidates = catalog_by_name.get(normalize_price_text(str(assortment.get("name") or "")), [])
            local = candidates[0] if len(candidates) == 1 else None
        api_group, api_category = _api_group(assortment)
        name = str(assortment.get("name") or "").strip()
        if not name:
            continue
        cash = _sale_price(assortment, cash_name)
        cashless = _sale_price(assortment, cashless_name)
        api_prices = {"common": (cash, cashless)} if cash is not None and cashless is not None else {}
        items_by_assortment[identifier] = ItemSummary(
            local.callback_id if local else 0,
            local.name if local else name,
            api_group if api_group != "Без группы" else (local.group_name if local else api_group),
            api_category if api_group != "Без группы" else (local.category_name if local else api_category),
            api_prices, local.code if local else code,
        )
    return InventoryReference(
        built_on or date.today().isoformat(), inventory_catalog_signature(catalog),
        store_ids, items_by_assortment,
    )


def save_inventory_reference(path, reference: InventoryReference) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 3,
        "built_on": reference.built_on,
        "catalog_signature": reference.catalog_signature,
        "store_ids": list(reference.store_ids),
        "items": {
            identifier: {
                "name": item.name,
                "group_name": item.group_name,
                "category_name": item.category_name,
                "code": item.code,
                "cash": str(item.warehouse_prices["common"][0]) if "common" in item.warehouse_prices else None,
                "cashless": str(item.warehouse_prices["common"][1]) if "common" in item.warehouse_prices else None,
            }
            for identifier, item in reference.items_by_assortment.items()
        },
    }
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    temporary.replace(path)


def load_inventory_reference(path) -> InventoryReference | None:
    if not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if payload.get("version") != 3:
            return None
        items = {
            identifier: ItemSummary(
                0, value["name"], value["group_name"], value["category_name"],
                {"common": (Decimal(value["cash"]), Decimal(value["cashless"]))}
                if value.get("cash") is not None and value.get("cashless") is not None else {},
                value.get("code", ""),
            )
            for identifier, value in payload["items"].items()
        }
        return InventoryReference(
            str(payload["built_on"]), str(payload.get("catalog_signature", "")),
            tuple(payload["store_ids"]), items,
        )
    except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError):
        return None


def build_inventory_snapshot(
    client: MoySkladClient,
    catalog: list[ItemSummary],
    reference: InventoryReference | None = None,
) -> InventorySnapshot:
    reference = reference or build_inventory_reference(client, catalog)
    store_ids = reference.store_ids
    positions: dict[str, list[Decimal]] = {}
    for row in client.current_availability(store_ids):
        assortment_id = str(row.get("assortmentId") or "").strip()
        store_id = str(row.get("storeId") or "").strip()
        if not assortment_id or store_id not in store_ids or assortment_id not in reference.items_by_assortment:
            continue
        if assortment_id not in positions:
            positions[assortment_id] = [Decimal(0)] * len(store_ids)
        try:
            quantity = max(Decimal(0), Decimal(str(row.get("quantity") or 0)))
        except Exception:
            continue
        positions[assortment_id][store_ids.index(store_id)] = quantity

    matched: dict[str, InventoryItem] = {}
    for identifier, local in reference.items_by_assortment.items():
        balances = tuple(positions.get(identifier, [Decimal(0)] * len(store_ids)))
        key = local.code.strip().casefold() or normalize_price_text(local.name)
        existing = matched.get(key)
        if existing:
            balances = tuple(existing.quantities[index] + balances[index] for index in range(len(store_ids)))
        matched[key] = InventoryItem(
            key, local.name, local.group_name, local.category_name, balances
        )

    items = tuple(sorted(matched.values(), key=lambda item: normalize_price_text(item.name)))
    grouped: dict[tuple[str, str], list[InventoryItem]] = {}
    for item in items:
        grouped.setdefault((item.group_name, item.category_name), []).append(item)
    groups = tuple(
        InventoryGroup(
            hashlib.sha256(f"{category}|{name}".encode()).hexdigest()[:16],
            name,
            category,
            tuple(values),
        )
        for (name, category), values in sorted(
            grouped.items(), key=lambda value: normalize_price_text(value[0][0])
        )
    )
    return InventorySnapshot(items, groups)


def inventory_cards(snapshot: InventorySnapshot, rules: dict[str, InventoryGroupingRule]) -> tuple[InventoryCard, ...]:
    cards: dict[str, dict[str, list[InventoryItem]]] = {}
    meta = {}
    for group in snapshot.groups:
        rule = rules.get(rule_key(group.name, group.category_name))
        if not rule or rule.excluded: continue
        cards.setdefault(rule.card_name, {}).setdefault(rule.line_name, []).extend(group.items)
        meta[(rule.card_name, rule.line_name)] = rule
    result = []
    for card, lines in cards.items():
        values = tuple(InventoryLine(name, meta[(card, name)].kind, meta[(card, name)].unit, tuple(items)) for name, items in lines.items())
        result.append(InventoryCard(hashlib.sha256(card.encode()).hexdigest()[:16], card, values))
    return tuple(sorted(result, key=lambda x: normalize_price_text(x.name)))


def search_inventory_cards(cards: tuple[InventoryCard, ...], query: str, limit: int = 8) -> list[InventoryCard]:
    found = []
    for card in cards:
        score = _score(query, card.name + " " + " ".join(line.name for line in card.lines))
        if score is not None: found.append((score, card))
    return [x for _, x in sorted(found, key=lambda x: x[0], reverse=True)[:limit]]


def _score(query: str, value: str) -> float | None:
    tokens = normalize_price_text(query).split()
    candidate = normalize_price_text(value)
    words = candidate.split()
    if not tokens:
        return None
    scores = []
    for token in tokens:
        if token in words:
            scores.append(1.0)
        elif token in candidate:
            scores.append(0.93)
        elif token.isdigit():
            scores.append(0.0)
        else:
            scores.append(max((SequenceMatcher(None, token, word).ratio() for word in words), default=0))
    if not all(value >= (0.82 if len(token) <= 2 else 0.66) for token, value in zip(tokens, scores)):
        return None
    return sum(scores) / len(scores) + (0.15 if normalize_price_text(query) in candidate else 0)


def search_inventory(
    snapshot: InventorySnapshot, query: str, *, group_limit: int = 8, item_limit: int = 12
) -> tuple[list[InventoryGroup], list[InventoryItem]]:
    groups = []
    for group in snapshot.groups:
        value = f"{group.name} {group.category_name} " + " ".join(item.name for item in group.items)
        score = _score(query, value)
        if score is not None:
            groups.append((score, group.total, group))
    groups.sort(key=lambda value: (value[0], value[1]), reverse=True)
    items = []
    for item in snapshot.items:
        score = _score(query, f"{item.name} {item.group_name} {item.category_name}")
        if score is not None:
            items.append((score, item.total, item))
    items.sort(key=lambda value: (value[0], value[1]), reverse=True)
    return (
        [group for _, _, group in groups[:group_limit]],
        [item for _, _, item in items[:item_limit]],
    )


def _quantity(value: Decimal) -> str:
    if value == value.to_integral_value():
        return f"{int(value):,}".replace(",", " ")
    return f"{value:,.3f}".rstrip("0").rstrip(".").replace(",", " ")


def _counted_word(value: int, forms: tuple[str, str, str]) -> str:
    remainder_100 = value % 100
    remainder_10 = value % 10
    if 11 <= remainder_100 <= 14:
        word = forms[2]
    elif remainder_10 == 1:
        word = forms[0]
    elif 2 <= remainder_10 <= 4:
        word = forms[1]
    else:
        word = forms[2]
    return f"{value} {word}"


def assortment_count(value: int, category_name: str) -> str:
    normalized = normalize_price_text(category_name)
    if any(marker in normalized for marker in ("жидкост", "конструктор", "ароматизатор", "однораз")):
        return _counted_word(value, ("вкус", "вкуса", "вкусов"))
    if any(marker in normalized for marker in ("электронные системы", "устройств", "желез")):
        return _counted_word(value, ("цвет", "цвета", "цветов"))
    return _counted_word(value, ("вариант", "варианта", "вариантов"))


def inventory_item_text(item: InventoryItem, updated_at: str) -> str:
    lines = [
        "📦 <b>Остатки товара</b>",
        "",
        f"<b>{html.escape(item.name)}</b>",
        f"<i>{html.escape(item.group_name)}</i>",
        "",
        f"<blockquote><b>{_quantity(item.total)} шт.</b> доступно на трёх складах</blockquote>",
        "",
        "🏙 <b>По городам</b>",
    ]
    lines.extend(
        f"• {html.escape(STORE_LABELS[name])} — <b>{_quantity(quantity)} шт.</b>"
        for name, quantity in zip(DEFAULT_STORES, item.quantities)
    )
    lines.extend(("", f"<blockquote>🕒 Актуально на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)


def inventory_group_text(group: InventoryGroup, updated_at: str) -> str:
    lines = [
        "📦 <b>Остатки товарной группы</b>",
        "",
        f"<b>{html.escape(group.name)}</b>",
        f"<i>{html.escape(group.category_name)}</i>",
        "",
        "<blockquote>"
        f"<b>{_quantity(group.total)} шт.</b> всего\n"
        f"<b>{assortment_count(len(group.items), group.category_name)}</b> в наличии"
        "</blockquote>",
        "",
        "🏙 <b>По городам</b>",
    ]
    for index, name in enumerate(DEFAULT_STORES):
        lines.append(
            f"• {html.escape(STORE_LABELS[name])} — "
            f"<b>{_quantity(group.quantities[index])} шт.</b> · "
            f"{assortment_count(group.variants_by_store[index], group.category_name)}"
        )
    lines.extend(("", f"<blockquote>🕒 Актуально на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)


def inventory_card_text(card: InventoryCard, updated_at: str) -> str:
    lines = ["📦 <b>Остатки по линейкам</b>", "", f"<b>{html.escape(card.name)}</b>"]
    for line in card.lines:
        lines.extend(("", f"<b>{html.escape(line.name)}</b>"))
        total_variants = len(line.items)
        for index, store in enumerate(DEFAULT_STORES):
            quantity = sum((item.quantities[index] for item in line.items), Decimal(0))
            available = sum(1 for item in line.items if item.quantities[index] > 0)
            noun = assortment_count(total_variants, line.kind).split(" ", 1)[1]
            lines.append(f"• {STORE_LABELS[store]} — <b>{_quantity(quantity)} {html.escape(line.unit)}</b> · {available}/{total_variants} {noun}")
    lines.extend(("", f"<blockquote>🕒 Актуально на {html.escape(updated_at)}</blockquote>"))
    return "\n".join(lines)
