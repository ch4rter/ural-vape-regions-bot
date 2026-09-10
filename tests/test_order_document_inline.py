from decimal import Decimal

from order_document_inline import (
    order_document_caption,
    search_customer_orders,
    select_print_template,
)


class FakeClient:
    def _rows(self, endpoint, params):
        if endpoint == "entity/counterparty":
            if params["search"] == "Vape Zone":
                return [{
                    "id": "c1", "name": "Vape Zone",
                    "meta": {"href": "https://api.moysklad.ru/api/remap/1.2/entity/counterparty/c1"},
                }]
            return []
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


class FuzzyFakeClient:
    def __init__(self):
        self.probes = []

    def _rows(self, endpoint, params):
        if endpoint == "entity/counterparty":
            self.probes.append(params["search"])
            if params["search"].casefold() in {"zone", "zon"}:
                return [{
                    "id": "c1",
                    "name": "(О) Vape Zone/Алексеев П.Б.-ОСНО/Санкт-Петербург",
                    "meta": {"href": "https://api.moysklad.ru/api/remap/1.2/entity/counterparty/c1"},
                }]
            return []
        return [{
            "id": "o1", "name": "100", "moment": "2026-09-09 12:00:00.000",
            "agent": {"name": "Vape Zone"}, "state": {"name": "Новый"},
            "organization": {"name": "ИП Селетков"}, "store": {"name": "Мордор"},
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


def test_select_print_template_falls_back_to_standard_customer_order():
    templates = [{
        "name": "Заказ покупателя",
        "meta": {"href": "https://api.moysklad.ru/api/remap/1.2/entity/customerorder/metadata/embeddedtemplate/1"},
    }]
    assert select_print_template(templates, "Накладная") == templates[0]


def test_search_customer_orders_accepts_incomplete_and_mistyped_name():
    client = FuzzyFakeClient()
    rows = search_customer_orders(client, "vepe zone")
    assert len(rows) == 1
    assert "zone" in [probe.casefold() for probe in client.probes]
