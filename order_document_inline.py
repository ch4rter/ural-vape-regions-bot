"""Live MoySklad customer-order lookup for inline PDF print forms."""

from __future__ import annotations

import html
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal

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


def search_customer_orders(
    client: MoySkladClient, query: str, *, limit: int = 6
) -> list[OrderDocument]:
    query = " ".join(query.strip().split())
    if len(query) < 3:
        return []
    counterparties = client._rows("entity/counterparty", {"limit": 5, "search": query})
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
    return partial[0] if len(partial) == 1 else None


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
