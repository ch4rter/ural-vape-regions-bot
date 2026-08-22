"""Read-only MoySklad customer-order status notification helpers."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from urllib.parse import urlparse

from materials_db import AccessUser, MaterialsDB
from moysklad_price import MoySkladClient, _canonical_href


@dataclass(frozen=True)
class OrderNotice:
    order_id: str
    name: str
    updated: str
    moment: str
    agent: str
    state_name: str
    state_href: str
    channel_name: str
    channel_href: str
    total: Decimal
    web_url: str


def _expanded_name(value: object, default: str = "—") -> str:
    if not isinstance(value, dict):
        return default
    return str(value.get("name") or default).strip()


def _expanded_href(value: object) -> str:
    if not isinstance(value, dict):
        return ""
    return _canonical_href(str(value.get("meta", {}).get("href", "")))


def parse_order_notice(row: dict) -> OrderNotice | None:
    order_id = str(row.get("id", "")).strip()
    state_href = _expanded_href(row.get("state"))
    state_name = _expanded_name(row.get("state"), "")
    if not order_id or not state_href or not state_name:
        return None
    try:
        total = Decimal(str(row.get("sum") or 0)) / Decimal(100)
    except Exception:
        total = Decimal(0)
    raw_url = str(row.get("meta", {}).get("uuidHref", "")).strip()
    parsed = urlparse(raw_url)
    web_url = raw_url if parsed.scheme == "https" and parsed.hostname == "online.moysklad.ru" else ""
    return OrderNotice(
        order_id=order_id,
        name=str(row.get("name") or "Без номера").strip(),
        updated=str(row.get("updated") or "").strip(),
        moment=str(row.get("moment") or "").strip(),
        agent=_expanded_name(row.get("agent")),
        state_name=state_name,
        state_href=state_href,
        channel_name=_expanded_name(row.get("salesChannel"), "Не указан"),
        channel_href=_expanded_href(row.get("salesChannel")),
        total=total,
        web_url=web_url,
    )


def fetch_changed_orders(client: MoySkladClient, updated_from: str) -> list[OrderNotice]:
    return [
        notice
        for row in client.changed_customer_orders(updated_from)
        if (notice := parse_order_notice(row)) is not None
    ]


def _valid_recipient(user: AccessUser | None) -> int | None:
    return user.telegram_id if user and user.telegram_id else None


def review_notification_targets(
    database: MaterialsDB, admin_ids: set[int]
) -> tuple[int, ...]:
    reviewer_raw = database.get_setting("order_review_reviewer_access_id") or ""
    reviewer = database.get_access_user(int(reviewer_raw)) if reviewer_raw.isdigit() else None
    candidates = [_valid_recipient(reviewer), *sorted(admin_ids)]
    return tuple(dict.fromkeys(value for value in candidates if value))


def checked_notification_targets(
    database: MaterialsDB,
    channel_href: str,
    admin_ids: set[int],
) -> tuple[int, ...]:
    manager = database.access_user_by_channel(channel_href) if channel_href else None
    assistant = database.team_assistant(manager.id) if manager else None
    # Delivery is deliberately sequential: assistant -> manager -> main admin.
    candidates = [
        _valid_recipient(assistant),
        _valid_recipient(manager),
        *sorted(admin_ids),
    ]
    return tuple(dict.fromkeys(value for value in candidates if value))
