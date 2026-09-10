"""Live MoySklad customer-order lookup for inline PDF print forms."""

from __future__ import annotations

import html
import re
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from difflib import SequenceMatcher

from moysklad_price import MoySkladClient


@dataclass(frozen=True)
class OrderDocument:
    order_id: str
    number: str
    moment: str
    updated: str
    agent: str
    state: str
    organization: str
    warehouse: str
    total: Decimal
    paid: Decimal


def _name(value: object, default: str = "—") -> str:
    return str(value.get("name") or default).strip() if isinstance(value, dict) else default


def _money(value: object) -> Decimal:
    try:
        return Decimal(str(value or 0)) / Decimal(100)
    except Exception:
        return Decimal(0)


def _search_text(value: object) -> str:
    text = str(value or "").casefold().replace("ё", "е")
    return " ".join(re.findall(r"[a-zа-я0-9]+", text))


def _counterparty_score(name: object, query: str) -> float:
    candidate = _search_text(name)
    needle = _search_text(query)
    if not candidate or not needle:
        return 0
    words = needle.split()
    matched = sum(
        1 for word in words
        if word in candidate or any(
            len(word) >= 4 and SequenceMatcher(None, word, part).ratio() >= 0.76
            for part in candidate.split()
        )
    )
    coverage = matched / len(words)
    phrase_bonus = 0.35 if needle in candidate else 0
    similarity = SequenceMatcher(None, needle, candidate).ratio() * 0.15
    return coverage + phrase_bonus + similarity


def _find_counterparties(client: MoySkladClient, query: str, *, limit: int = 5) -> list[dict]:
    """Use a few broad API searches and rank the results locally.

    MoySklad's ``search`` is not fuzzy and treats a multi-word phrase quite
    strictly. Searching by distinctive word prefixes lets employees use a
    shortened name or make a small typo without downloading the whole client
    directory.
    """
    normalized = _search_text(query)
    words = [word for word in normalized.split() if len(word) >= 3]
    probes = [query.strip()]
    for word in sorted(words, key=len, reverse=True):
        probes.extend((word, word[: max(4, len(word) - 1)]))

    candidates: dict[str, dict] = {}
    for probe in dict.fromkeys(value for value in probes if len(value) >= 3):
        for row in client._rows("entity/counterparty", {"limit": 25, "search": probe}):
            key = str(row.get("id") or row.get("meta", {}).get("href") or "")
            if key:
                candidates[key] = row
        if len(candidates) >= 25:
            break

    ranked = sorted(
        candidates.values(),
        key=lambda row: _counterparty_score(row.get("name"), query),
        reverse=True,
    )
    return [row for row in ranked if _counterparty_score(row.get("name"), query) >= 0.45][:limit]


def search_customer_orders(
    client: MoySkladClient, query: str, *, limit: int = 6
) -> list[OrderDocument]:
    query = " ".join(query.strip().split())
    if len(query) < 3:
        return []
    counterparties = _find_counterparties(client, query, limit=5)
    orders = []
    for counterparty in counterparties:
        href = str(counterparty.get("meta", {}).get("href") or "")
        if not href:
            continue
        rows = client._rows("entity/customerorder", {
            "limit": limit,
            "filter": f"agent={href}",
            "order": "moment,desc",
            "expand": "agent,state,organization,store",
        })
        orders.extend(rows)
    orders.sort(key=lambda row: str(row.get("moment") or ""), reverse=True)
    unique_orders = {str(row.get("id") or ""): row for row in orders if row.get("id")}
    orders = sorted(
        unique_orders.values(), key=lambda row: str(row.get("moment") or ""), reverse=True
    )
    return [
        OrderDocument(
            order_id=str(row.get("id") or ""),
            number=str(row.get("name") or "Без номера").strip(),
            moment=str(row.get("moment") or ""),
            updated=str(row.get("updated") or ""),
            agent=_name(row.get("agent")),
            state=_name(row.get("state")),
            organization=_name(row.get("organization")),
            warehouse=_name(row.get("store")),
            total=_money(row.get("sum")),
            paid=_money(row.get("payedSum")),
        )
        for row in orders[:limit]
        if row.get("id")
    ]


def select_print_template(templates: list[dict], configured_name: str = "Накладная") -> dict | None:
    named = [(str(value.get("name") or "").strip(), value) for value in templates]
    needle = configured_name.strip().casefold()
    exact = [value for name, value in named if name.casefold() == needle]
    if exact:
        return exact[0]
    partial = [value for name, value in named if needle and needle in name.casefold()]
    if len(partial) == 1:
        return partial[0]
    # A fresh MoySklad account often has only the standard customer-order form.
    # It is the safe built-in equivalent when no custom "Накладная" exists.
    if needle == "накладная":
        order_forms = [
            value for name, value in named
            if name.casefold() in {"заказ покупателя", "заказ покупателя (счет)"}
        ]
        if order_forms:
            return order_forms[0]
    return None


def rubles(value: Decimal) -> str:
    return f"{value:,.2f}".replace(",", " ").replace(".00", "").replace(".", ",") + " ₽"


def order_document_description(order: OrderDocument) -> str:
    return f"{order.state} · {rubles(order.total)} · {order.warehouse}"


def order_document_caption(order: OrderDocument) -> str:
    try:
        moment = datetime.fromisoformat(order.moment[:19]).strftime("%d.%m.%Y")
    except ValueError:
        moment = order.moment[:10] or "—"
    return "\n".join((
        f"📄 <b>Накладная к заказу №{html.escape(order.number)}</b>",
        "",
        f"<b>Контрагент:</b> {html.escape(order.agent)}",
        f"<b>Дата:</b> {html.escape(moment)}",
        f"<b>Статус:</b> {html.escape(order.state)}",
        f"<b>Сумма:</b> {rubles(order.total)}",
        f"<b>Оплачено:</b> {rubles(order.paid)}",
        f"<b>Склад:</b> {html.escape(order.warehouse)}",
        f"<b>Организация:</b> {html.escape(order.organization)}",
    ))
