from decimal import Decimal

from order_document_inline import (
    order_document_caption,
    search_customer_orders,
    select_print_template,
)


class FakeClient:
    def _rows(self, endpoint, params):
        if endpoint == "entity/counterparty":
            assert params["search"] == "Vape Zone"
            return [{"meta": {"href": "https://api.moysklad.ru/api/remap/1.2/entity/counterparty/c1"}}]
        assert endpoint == "entity/customerorder"
        assert "counterparty/c1" in params["filter"]
        return [{
            "id": "o1", "name": "04262", "moment": "2026-09-08 12:30:00.000",
            "updated": "2026-09-09 09:00:00.000", "sum": 8640000, "payedSum": 5000000,
            "agent": {"name": "Vape Zone / ИП Алексеев"},
            "state": {"name": "Проверено"},
            "organization": {"name": "ИП Селетков В.Г."},
            "store": {"name": "Мордор"},
        }]


def test_search_customer_orders_and_caption():
    rows = search_customer_orders(FakeClient(), "Vape Zone")
    assert len(rows) == 1
    order = rows[0]
    assert order.total == Decimal("86400")
    assert order.paid == Decimal("50000")
    text = order_document_caption(order)
    assert "заказу №04262" in text
    assert "Vape Zone / ИП Алексеев" in text
    assert "86 400 ₽" in text
    assert "ИП Селетков В.Г." in text


def test_select_print_template_requires_unambiguous_match():
    templates = [
        {"name": "Заказ покупателя", "meta": {"href": "order"}},
        {"name": "Накладная", "meta": {"href": "waybill"}},
    ]
    assert select_print_template(templates, "накладная")["meta"]["href"] == "waybill"
    assert select_print_template(templates, "несуществующая") is None
