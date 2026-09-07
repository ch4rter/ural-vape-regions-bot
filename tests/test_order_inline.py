from decimal import Decimal
from datetime import datetime
from io import BytesIO
from pathlib import Path

from PIL import Image

from materials_db import MaterialsDB
from order_inline import fetch_inline_order, format_order_caption, render_order_card
from prices_db import ItemSummary


class FakeClient:
    def _rows(self, endpoint, params):
        assert endpoint == "entity/customerorder"
        assert params["search"] == "18452"
        return [{
            "id": "order-1",
            "name": "18452",
            "updated": "2026-09-07 12:00:00.000",
            "moment": "2026-09-07 11:42:00.000",
            "sum": 18645000,
            "payedSum": 12000000,
            "agent": {"name": "Vape Zone"},
            "state": {"name": "Ожидает проверки"},
            "salesChannel": {"name": "Валера"},
            "organization": {"name": "ИП Иванов И.И."},
            "store": {"name": "Мордор"},
            "positions": {"meta": {"href": "https://api.moysklad.ru/positions"}},
        }]

    def _get_json(self, url, params=None):
        assert params == {"limit": 1000, "expand": "assortment"}
        return {"rows": [
            {
                "quantity": 10, "price": 1000000, "discount": 0,
                "assortment": {"name": "OGGO VLIQ Ice"},
            },
            {
                "quantity": 5, "price": 500000, "discount": 10,
                "assortment": {"name": "Vaporesso XROS"},
            },
        ], "meta": {}}


def test_fetch_and_render_inline_order_card():
    items = [
        ItemSummary(1, "OGGO VLIQ Ice", "OGGO VLIQ", "Жидкости", {}),
        ItemSummary(2, "Vaporesso XROS", "Vaporesso устройства", "Устройства", {}),
    ]
    order = fetch_inline_order(FakeClient(), "18452", items)
    assert order is not None
    assert order.total == Decimal("186450")
    assert order.paid == Decimal("120000")
    assert order.organization == "ИП Иванов И.И."
    assert order.warehouse == "Мордор"
    assert order.groups == (
        ("OGGO VLIQ", Decimal("100000")),
        ("Vaporesso устройства", Decimal("22500.0")),
    )
    root = Path(__file__).resolve().parents[1]
    content = render_order_card(
        order,
        root / "assets/order_card_background.png",
        root / "assets/ural_vape_logo_transparent.png",
    )
    image = Image.open(BytesIO(content))
    assert image.size == (1280, 1280)
    assert image.format == "JPEG"
    assert len(content) < 2_000_000


def test_inline_card_file_id_cache(tmp_path):
    database = MaterialsDB(tmp_path / "materials.sqlite3")
    assert database.inline_order_card_file("order-1", "fingerprint") is None
    database.save_inline_order_card_file("order-1", "fingerprint", "telegram-file")
    assert database.inline_order_card_file("order-1", "fingerprint") == "telegram-file"
    assert database.inline_order_card_file("order-1", "changed") is None


def test_caption_adds_history_without_repeating_current_order():
    order = fetch_inline_order(FakeClient(), "18452", [])
    caption = format_order_caption(order, {
        "last_moment": datetime(2026, 8, 1, 10, 0),
        "days_since_last": 10,
        "recent_purchases": 3,
        "recent_revenue": 150000,
        "average_purchase": 50000,
        "expected_days": 25,
        "history_shipments": 3,
    }, now=datetime(2026, 8, 11, 12, 0))
    assert "За 90 дней" in caption
    assert "50 000 ₽" in caption
    assert "больше средней" in caption
    assert "Заказ №" not in caption
    assert "Ожидает проверки" not in caption
