"""Read-only MoySklad order lookup and branded Telegram inline cards."""

from __future__ import annotations

import hashlib
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

CARD_TEMPLATE_VERSION = "2"


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
    total: Decimal
    paid: Decimal
    groups: tuple[tuple[str, Decimal], ...]

    @property
    def fingerprint(self) -> str:
        source = "|".join((
            CARD_TEMPLATE_VERSION, self.order_id, self.updated, self.organization,
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
        "expand": "agent,state,salesChannel,organization",
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

    client_font = _fit(draw, order.agent, 1040, 54, 32)
    draw.text((640, 330), order.agent, font=client_font, fill=white, anchor="ma")
    draw.text((640, 397), f"Менеджер · {order.channel}", font=_font(26), fill=muted, anchor="ma")
    organization = f"Организация · {order.organization}"
    organization_font = _fit(draw, organization, 1030, 25, 19)
    draw.text((640, 438), organization, font=organization_font, fill=muted, anchor="ma")

    cards = (("ЗАКАЗАНО", order.total), ("ОПЛАЧЕНО", order.paid), ("ОСТАЛОСЬ", max(Decimal(0), order.total - order.paid)))
    card_left, card_width, card_gap = 105, 340, 25
    for index, (label, amount) in enumerate(cards):
        left = card_left + index * (card_width + card_gap)
        draw.rounded_rectangle((left, 492, left + card_width, 632), radius=25, fill=(255, 255, 255, 18), outline=(255, 255, 255, 34), width=2)
        draw.text((left + card_width // 2, 526), label, font=_font(22, bold=True), fill=muted, anchor="mm")
        amount_text = _rubles(amount)
        draw.text((left + card_width // 2, 579), amount_text, font=_fit(draw, amount_text, 295, 37, 25), fill=white, anchor="mm")

    draw.text((105, 687), "СОСТАВ ЗАКАЗА", font=_font(24, bold=True), fill=accent)
    shown = list(order.groups[:4])
    if len(order.groups) > 4:
        shown[-1] = (f"Ещё {len(order.groups) - 3} групп", sum(value for _, value in order.groups[3:]))
    if not shown:
        shown = [("Состав пока не классифицирован", order.total)]
    y = 742
    for name, amount in shown:
        clean_name = " ".join(name.split())
        if len(clean_name) > 40:
            clean_name = textwrap.shorten(clean_name, width=40, placeholder="…")
        draw.text((118, y), clean_name, font=_fit(draw, clean_name, 760, 28, 21), fill=white)
        amount_text = _rubles(amount)
        amount_width = draw.textbbox((0, 0), amount_text, font=_font(28, bold=True))[2]
        draw.text((1162 - amount_width, y), amount_text, font=_font(28, bold=True), fill=white)
        y += 57

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
