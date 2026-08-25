from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from moysklad_bonus import month_bounds
from moysklad_price import MoySkladClient


@dataclass(frozen=True)
class ManagerProgress:
    name: str
    channel_href: str
    revenue: Decimal


def canonical_href(value: str) -> str:
    return value.split("?", 1)[0].rstrip("/")


def parse_goal(value: str | None) -> Decimal | None:
    if not value:
        return None
    try:
        result = Decimal(value)
    except (InvalidOperation, ValueError):
        return None
    return result if result > 0 else None


def fetch_team_revenue(
    client: MoySkladClient,
    month: str,
    managers: list[tuple[str, str]],
) -> list[ManagerProgress]:
    """Read conducted shipments once and aggregate only configured channels."""
    start, end = month_bounds(month)
    documents = client._rows(
        "entity/demand",
        {
            "limit": 100,
            "filter": f"moment>={start};moment<{end};applicable=true",
            "expand": "salesChannel",
        },
    )
    totals: dict[str, Decimal] = {}
    for document in documents:
        channel = document.get("salesChannel")
        href = canonical_href(str(channel.get("meta", {}).get("href", ""))) if isinstance(channel, dict) else ""
        if not href:
            continue
        try:
            amount = Decimal(str(document.get("sum") or 0)) / Decimal(100)
        except (InvalidOperation, ValueError):
            continue
        totals[href] = totals.get(href, Decimal(0)) + amount

    rows = [
        ManagerProgress(name, canonical_href(href), totals.get(canonical_href(href), Decimal(0)))
        for name, href in managers
    ]
    return sorted(rows, key=lambda row: (-row.revenue, row.name.casefold()))


def format_money(value: Decimal) -> str:
    value = value.quantize(Decimal("0.01"))
    if value == value.to_integral_value():
        return f"{int(value):,}".replace(",", " ") + " ₽"
    return f"{value:,.2f}".replace(",", " ").replace(".", ",") + " ₽"


def progress_bar(total: Decimal, goal: Decimal, size: int = 10) -> str:
    ratio = max(Decimal(0), total / goal) if goal else Decimal(0)
    filled = min(size, int(ratio * size))
    return "🟩" * filled + "⬜" * (size - filled)


def month_timing(month: str, today: date) -> tuple[str, Decimal | None]:
    year, month_number = map(int, month.split("-"))
    first = date(year, month_number, 1)
    last = date(year, month_number, calendar.monthrange(year, month_number)[1])
    if today < first:
        return "📅 Месяц ещё не начался", None
    if today > last:
        return "📅 Месяц завершён", None
    days = (last - today).days + 1
    return f"📅 До конца месяца: <b>{days}</b>", Decimal(days)


def render_progress(
    month_label: str,
    month: str,
    goal: Decimal,
    managers: list[ManagerProgress],
    *,
    now: datetime,
) -> str:
    managers = sorted(managers, key=lambda row: (-row.revenue, row.name.casefold()))
    total = sum((row.revenue for row in managers), Decimal(0))
    remaining = max(goal - total, Decimal(0))
    percent = total / goal * Decimal(100) if goal else Decimal(0)
    timing, days = month_timing(month, now.date())
    medals = ("🥇", "🥈", "🥉")
    ranking = []
    for index, row in enumerate(managers):
        marker = medals[index] if index < len(medals) else f"{index + 1}."
        ranking.append(f"{marker} <b>{row.name}</b> — {format_money(row.revenue)}")

    percent_text = str(percent.quantize(Decimal("0.1"))).replace(".", ",")
    lines = [
        f"{progress_bar(total, goal)}  <b>{percent_text}%</b>",
        "",
        f"🔥 <b>ЦЕЛЬ ОТДЕЛА · {month_label.upper()}</b>",
        "",
        f"💰 <b>Выполнено:</b> {format_money(total)}",
        f"🏆 <b>Цель:</b> {format_money(goal)}",
        f"🎯 <b>Осталось:</b> {format_money(remaining)}",
        "",
        "👥 <b>Вклад команды</b>",
        "",
        *ranking,
        "",
        timing,
    ]
    if days and remaining:
        lines.append(f"⚡ <b>Для выполнения:</b> {format_money(remaining / days)} в день")
    if total >= goal:
        lines.extend(["", "🏆 <b>ЦЕЛЬ ВЫПОЛНЕНА!</b>"])
    lines.extend(["", f"🕒 Обновлено: {now:%d.%m.%Y в %H:%M:%S}"])
    return "\n".join(lines)
