from decimal import Decimal

from inventory_inline import (
    build_inventory_reference,
    build_inventory_snapshot,
    inventory_group_text,
    inventory_item_text,
    load_inventory_reference,
    save_inventory_reference,
    search_inventory,
)
from prices_db import ItemSummary


class FakeInventoryClient:
    def stores(self):
        return [
            {"name": name, "meta": {"href": f"https://api.moysklad.ru/entity/store/{index}"}}
            for index, name in enumerate(("Мордор", "Годзибасы", "Жможики"), 1)
        ]

    def current_availability(self, store_ids):
        assert store_ids == ("1", "2", "3")
        return [
            {"assortmentId": "a", "storeId": "1", "quantity": 12},
            {"assortmentId": "a", "storeId": "2", "quantity": 3},
            {"assortmentId": "b", "storeId": "2", "quantity": 5},
            {"assortmentId": "b", "storeId": "3", "quantity": 7},
            {"assortmentId": "a", "storeId": "technical", "quantity": 999},
        ]

    def assortment(self):
        return [
            {
                "name": "Картридж XROS 0.6 2 мл", "code": "001",
                "meta": {"href": "https://api.moysklad.ru/entity/product/a"},
            },
            {
                "name": "Картридж XROS 0.8 2 мл", "code": "002",
                "meta": {"href": "https://api.moysklad.ru/entity/product/b"},
            },
        ]


def test_inventory_snapshot_search_and_public_labels():
    catalog = [
        ItemSummary(1, "Картридж XROS 0.6 2 мл", "Vaporesso расходники", "Картриджи", {}, "001"),
        ItemSummary(2, "Картридж XROS 0.8 2 мл", "Vaporesso расходники", "Картриджи", {}, "002"),
    ]
    snapshot = build_inventory_snapshot(FakeInventoryClient(), catalog)
    assert len(snapshot.items) == 2
    group = snapshot.groups[0]
    assert group.total == Decimal("27")
    assert group.quantities == (Decimal("12"), Decimal("8"), Decimal("7"))
    assert group.variants_by_store == (1, 2, 1)

    groups, items = search_inventory(snapshot, "картриджи xros")
    assert groups[0].name == "Vaporesso расходники"
    assert len(items) == 2
    group_text = inventory_group_text(groups[0], "08.09.2026 15:00")
    item_text = inventory_item_text(items[0], "08.09.2026 15:00")
    assert "Москва" in group_text
    assert "Санкт-Петербург" in group_text
    assert "Урал" in group_text
    assert "Мордор" not in group_text
    assert "₽" not in group_text + item_text


def test_daily_reference_can_be_reused_without_assortment_requests(tmp_path):
    catalog = [
        ItemSummary(1, "Картридж XROS 0.6 2 мл", "Vaporesso расходники", "Картриджи", {}, "001"),
        ItemSummary(2, "Картридж XROS 0.8 2 мл", "Vaporesso расходники", "Картриджи", {}, "002"),
    ]
    client = FakeInventoryClient()
    reference = build_inventory_reference(client, catalog, built_on="2026-09-08")
    path = tmp_path / "reference.json"
    save_inventory_reference(path, reference)
    restored = load_inventory_reference(path)
    assert restored is not None
    assert restored.built_on == "2026-09-08"
    client.stores = lambda: (_ for _ in ()).throw(AssertionError("stores reread"))
    client.assortment = lambda: (_ for _ in ()).throw(AssertionError("assortment reread"))
    snapshot = build_inventory_snapshot(client, catalog, restored)
    assert snapshot.groups[0].total == Decimal("27")
