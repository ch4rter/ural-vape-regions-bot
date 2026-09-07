"""Read-only MoySklad order lookup and branded Telegram inline cards."""

from __future__ import annotations

import hashlib
import html
import io
import re
import textwrap
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from urllib.parse import urlparse

from PIL import Image, ImageDraw, ImageFont

from moysklad_price import MoySkladClient, _canonical_href
from prices_db import ItemSummary, normalize_price_text

CARD_TEMPLATE_VERSION = "3"


@dataclass(frozen=True)
class InlineOrderCard:
    order_id: str
    name: str
    updated: str
    moment: str
    agent: str
    state: str
    channel: str
    organization: str
    warehouse: str
    total: Decimal
    paid: Decimal
    groups: tuple[tuple[str, Decimal], ...]

    @property
    def fingerprint(self) -> str:
        source = "|".join((
            CARD_TEMPLATE_VERSION, self.order_id, self.updated, self.organization,
            self.warehouse,
            str(self.total), str(self.paid), repr(self.groups),
        ))
        return hashlib.sha256(source.encode("utf-8")).hexdigest()


def _money(value: object) -> Decimal:
    try:
        return Decimal(str(value or 0)) / Decimal(100)
    except Exception:
        return Decimal(0)


def _name(value: object, default: str = "—") -> str:
    if not isinstance(value, dict):
        return default
    return str(value.get("name") or default).strip()


def _position_sum(position: dict) -> Decimal:
    quantity = Decimal(str(position.get("quantity") or 0))
    price = _money(position.get("price"))
    discount = Decimal(str(position.get("discount") or 0))
    return price * quantity * (Decimal(1) - discount / Decimal(100))


def _positions(client: MoySkladClient, order: dict) -> list[dict]:
    positions = order.get("positions") if isinstance(order.get("positions"), dict) else {}
    meta = positions.get("meta") if isinstance(positions.get("meta"), dict) else {}
    url = str(meta.get("href") or "")
    if not url:
        return [row for row in positions.get("rows", []) if isinstance(row, dict)]
    rows: list[dict] = []
    params = {"limit": 1000, "expand": "assortment"}
    while url:
        payload = client._get_json(url, params)
        params = None
        rows.extend(row for row in payload.get("rows", []) if isinstance(row, dict))
        url = str(payload.get("meta", {}).get("nextHref") or "")
    return rows


def _item_group_lookup(items: list[ItemSummary]) -> dict[str, str]:
    result = {}
    for item in items:
        key = normalize_price_text(item.name)
        if key:
            result[key] = item.group_name or item.category_name or "Прочее"
    return result


def fetch_inline_order(
    client: MoySkladClient,
    query: str,
    price_items: list[ItemSummary],
) -> InlineOrderCard | None:
    query = query.strip()
    if not re.fullmatch(r"[A-Za-zА-Яа-яЁё0-9._/-]{3,40}", query):
        return None
    rows = client._rows("entity/customerorder", {
        "limit": 20,
        "search": query,
        "expand": "agent,state,salesChannel,organization,store",
    })
    if not rows:
        return None
    normalized = query.casefold().lstrip("№#")
    exact = [row for row in rows if str(row.get("name") or "").casefold().lstrip("№#") == normalized]
    if not exact and len(rows) != 1:
        return None
    order = (exact or rows)[0]
    lookup = _item_group_lookup(price_items)
    group_values: dict[str, Decimal] = defaultdict(Decimal)
    for position in _positions(client, order):
        assortment = position.get("assortment") if isinstance(position.get("assortment"), dict) else {}
        item_name = str(assortment.get("name") or "").strip()
        path_name = str(assortment.get("pathName") or "").strip(" /\t")
        group = lookup.get(normalize_price_text(item_name))
        if not group and path_name:
            group = path_name.split("/")[-1].strip()
        group_values[group or "Прочее"] += _position_sum(position)
    groups = tuple(sorted(group_values.items(), key=lambda item: item[1], reverse=True))
    return InlineOrderCard(
        order_id=str(order.get("id") or ""),
        name=str(order.get("name") or "Без номера").strip(),
        updated=str(order.get("updated") or "").strip(),
        moment=str(order.get("moment") or "").strip(),
        agent=_name(order.get("agent")),
        state=_name(order.get("state")),
        channel=_name(order.get("salesChannel"), "Не указан"),
        organization=_name(order.get("organization"), "Не указана"),
        warehouse=_name(order.get("store"), "Не указан"),
        total=_money(order.get("sum")),
        paid=_money(order.get("payedSum")),
        groups=groups,
    )


def _font(size: int, *, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = (
        ("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ("/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf"),
    )
    for path in candidates:
        if Path(path).is_file():
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def _fit(draw: ImageDraw.ImageDraw, value: str, width: int, size: int, minimum: int = 24) -> ImageFont.FreeTypeFont:
    while size > minimum:
        font = _font(size, bold=True)
        if draw.textbbox((0, 0), value, font=font)[2] <= width:
            return font
        size -= 2
    return _font(minimum, bold=True)


def _rubles(value: Decimal) -> str:
    return f"{value:,.0f} ₽".replace(",", " ")


def format_order_caption(
    order: InlineOrderCard, history: dict, *, now: datetime | None = None
) -> str:
    """Complement the image with client history without repeating card fields."""
    if not history:
        return (
            "🆕 <b>Новый клиент</b>\n\n"
            "Проведённых отгрузок в доступной истории пока не найдено."
        )
    current = now or datetime.now()
    last_moment = history.get("last_moment")
    last_text = last_moment.strftime("%d.%m.%Y") if isinstance(last_moment, datetime) else "—"
    days = int(history.get("days_since_last") or 0)
    purchases = int(history.get("recent_purchases") or 0)
    revenue = Decimal(str(history.get("recent_revenue") or 0))
    average = Decimal(str(history.get("average_purchase") or 0))
    expected = history.get("expected_days")
    lines = [
        "📊 <b>История клиента</b>",
        "",
        f"Последняя отгрузка — <b>{html.escape(last_text)}</b> ({days} дн. назад)",
        f"За 90 дней — <b>{purchases}</b> закупок на <b>{_rubles(revenue)}</b>",
    ]
    if purchases:
        lines.append(f"Средняя закупка — <b>{_rubles(average)}</b>")
    if expected:
        lines.append(f"Обычный ритм — примерно раз в <b>{int(expected)} дней</b>")
    insight = "Клиент заказывает в своём обычном ритме."
    if expected and days > int(expected) * 1.5:
        insight = "🔄 Этот заказ оформлен после длительного перерыва."
    elif average and order.total >= average * Decimal("1.3"):
        change = ((order.total / average) - Decimal(1)) * Decimal(100)
        insight = f"🔥 Текущий заказ на <b>{change:.0f}%</b> больше средней закупки клиента."
    elif average and order.total <= average * Decimal("0.7"):
        change = (Decimal(1) - (order.total / average)) * Decimal(100)
        insight = f"Текущий заказ на <b>{change:.0f}%</b> меньше средней закупки клиента."
    lines.extend(("", insight, "", f"Обновлено · {current:%d.%m.%Y %H:%M}"))
    return "\n".join(lines)


def _status_color(status: str) -> tuple[int, int, int]:
    value = status.casefold()
    if "ожида" in value:
        return (245, 158, 11)
    if "отмен" in value or "проср" in value:
        return (220, 63, 75)
    if "выполн" in value or "закры" in value:
        return (110, 118, 128)
    if "провер" in value:
        return (43, 174, 102)
    return (43, 174, 102)


def _wrapped_heading(
    draw: ImageDraw.ImageDraw, value: str, width: int
) -> tuple[list[str], ImageFont.FreeTypeFont]:
    words = value.split()
    for size in range(52, 31, -2):
        font = _font(size, bold=True)
        if draw.textbbox((0, 0), value, font=font)[2] <= width:
            return [value], font
        choices = []
        for split_at in range(1, len(words)):
            lines = [" ".join(words[:split_at]), " ".join(words[split_at:])]
            widest = max(draw.textbbox((0, 0), line, font=font)[2] for line in lines)
            if widest <= width:
                choices.append((abs(len(lines[0]) - len(lines[1])), lines))
        if choices:
            return min(choices, key=lambda item: item[0])[1], font
    return [textwrap.shorten(value, width=62, placeholder="…")], _font(30, bold=True)


def render_order_card(
    order: InlineOrderCard,
    background_path: Path,
    logo_path: Path,
) -> bytes:
    canvas = Image.open(background_path).convert("RGB").resize((1280, 1280), Image.Resampling.LANCZOS)
    draw = ImageDraw.Draw(canvas, "RGBA")
    panel_left, panel_right = 75, 1205
    draw.rounded_rectangle(
        (panel_left, 45, panel_right, 1040), radius=42,
        fill=(15, 14, 23, 158), outline=(255, 255, 255, 32), width=2,
    )

    logo = Image.open(logo_path).convert("RGBA")
    logo.thumbnail((245, 122), Image.Resampling.LANCZOS)
    logo_x = (1280 - logo.width) // 2
    canvas.paste(logo, (logo_x, 72), logo)

    white = (255, 255, 255, 255)
    muted = (202, 195, 218, 255)
    accent = (237, 124, 49, 255)
    draw.text((105, 222), f"ЗАКАЗ №{order.name}", font=_font(31, bold=True), fill=muted)
    status = order.state.upper()
    status_rect = (735, 207, 1175, 269)
    status_font = _fit(draw, status, status_rect[2] - status_rect[0] - 42, 29, 20)
    draw.rounded_rectangle(status_rect, radius=27, fill=(*_status_color(order.state), 235))
    draw.text(((status_rect[0] + status_rect[2]) // 2, (status_rect[1] + status_rect[3]) // 2), status, font=status_font, fill=white, anchor="mm")

    client_lines, client_font = _wrapped_heading(draw, order.agent, 1050)
    client_y = 320 if len(client_lines) == 1 else 298
    for line in client_lines:
        draw.text((640, client_y), line, font=client_font, fill=white, anchor="ma")
        client_y += 53
    details_y = 400 if len(client_lines) == 1 else 418
    draw.text((640, details_y), f"Менеджер · {order.channel}", font=_font(25), fill=muted, anchor="ma")
    organization = f"Организация · {order.organization}"
    organization_font = _fit(draw, organization, 1030, 25, 19)
    draw.text((640, details_y + 39), organization, font=organization_font, fill=muted, anchor="ma")
    warehouse = f"Склад · {order.warehouse}"
    draw.text((640, details_y + 77), warehouse, font=_font(24), fill=muted, anchor="ma")

    cards = (("СУММА ЗАКАЗА", order.total), ("ОПЛАЧЕНО", order.paid))
    card_left, card_width, card_gap = 105, 522, 26
    for index, (label, amount) in enumerate(cards):
        left = card_left + index * (card_width + card_gap)
        draw.rounded_rectangle((left, 530, left + card_width, 670), radius=25, fill=(255, 255, 255, 18), outline=(255, 255, 255, 34), width=2)
        draw.text((left + card_width // 2, 564), label, font=_font(22, bold=True), fill=muted, anchor="mm")
        amount_text = _rubles(amount)
        draw.text((left + card_width // 2, 617), amount_text, font=_fit(draw, amount_text, 440, 39, 27), fill=white, anchor="mm")

    draw.text((105, 718), "СОСТАВ ЗАКАЗА", font=_font(24, bold=True), fill=accent)
    shown = list(order.groups[:5])
    if not shown:
        shown = [("Состав пока не классифицирован", order.total)]
    y = 768
    for name, amount in shown:
        clean_name = " ".join(name.split())
        if len(clean_name) > 40:
            clean_name = textwrap.shorten(clean_name, width=40, placeholder="…")
        draw.text((118, y), clean_name, font=_fit(draw, clean_name, 760, 28, 21), fill=white)
        amount_text = _rubles(amount)
        amount_width = draw.textbbox((0, 0), amount_text, font=_font(28, bold=True))[2]
        draw.text((1162 - amount_width, y), amount_text, font=_font(28, bold=True), fill=white)
        y += 46
    if len(order.groups) > len(shown):
        draw.text((118, y + 2), f"+ ещё {len(order.groups) - len(shown)} групп", font=_font(21), fill=muted)

    try:
        moment = datetime.fromisoformat(order.moment[:19]).strftime("%d.%m.%Y %H:%M")
    except ValueError:
        moment = order.moment[:16].replace("T", " ") or "не указана"
    draw.text((640, 1000), f"Создан · {moment}", font=_font(22), fill=(215, 209, 225, 230), anchor="mm")
    output = io.BytesIO()
    canvas.save(output, format="JPEG", quality=91, optimize=True)
    return output.getvalue()


def safe_moysklad_url(value: str) -> str:
    parsed = urlparse(value)
    return value if parsed.scheme == "https" and parsed.hostname == "online.moysklad.ru" else ""
