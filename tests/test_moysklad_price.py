from decimal import Decimal

from openpyxl import load_workbook

from moysklad_price import MoySkladError, build_price_from_moysklad
from prices_db import parse_price_file


class FakeMoySkladClient:
    def stores(self):
        return [
            {"name": name, "meta": {"href": f"https://api.moysklad.ru/api/remap/1.2/entity/store/{index}"}}
            for index, name in enumerate(("Мордор", "Годзибасы", "Жможики", "Другой"), 1)
        ]

    def current_availability(self, store_ids):
        assert store_ids == ("1", "2", "3")
        return [
            {
                "assortmentId": "available", "storeId": "1", "quantity": 7,
            },
            {
                "assortmentId": "available", "storeId": "3", "quantity": 2,
            },
            {
                "assortmentId": "reserved", "storeId": "1", "quantity": 0,
            },
            {
                "assortmentId": "other", "storeId": "4", "quantity": 20,
            },
        ]

    def assortment(self):
        prices = [
            {"priceType": {"name": "от 50т.р. нал"}, "value": 12345},
            {"priceType": {"name": "от 50т.р. безнал"}, "value": 13579},
        ]
        return [
            {
                "meta": {
                    "href": "https://api.moysklad.ru/api/remap/1.2/entity/product/available",
                    "type": "product",
                },
                "name": "OGGO VLIQ Вишня",
                "pathName": "ЭС/Жидкости/OGGO VLIQ",
                "salePrices": prices,
            },
            {
                "meta": {
                    "href": "https://api.moysklad.ru/api/remap/1.2/entity/product/reserved",
                    "type": "product",
                },
                "name": "Полностью зарезервирован",
                "pathName": "ЭС/Жидкости/OGGO VLIQ",
                "salePrices": prices,
            },
        ]


def test_builds_read_only_price_for_available_selected_stock(tmp_path):
    destination = tmp_path / "common.xlsx"
    result = build_price_from_moysklad(
        "test-token",
        destination,
        client=FakeMoySkladClient(),
    )

    assert result.item_count == 1
    assert result.group_count == 1
    parsed = parse_price_file(destination)
    assert parsed.item_count == 1
    item = parsed.groups[0].items[0]
    assert item.name == "OGGO VLIQ Вишня"
    assert item.cash == Decimal("123.45")
    assert item.cashless == Decimal("135.79")

    workbook = load_workbook(destination, data_only=True)
    sheet = workbook.active
    item_row = next(row for row in sheet.iter_rows(values_only=True) if row[0] == item.name)
    assert item_row[3:6] == (7, 0, 2)
    assert sheet.freeze_panes == "A4"
    workbook.close()


def test_fails_if_required_store_is_missing(tmp_path):
    client = FakeMoySkladClient()
    client.stores = lambda: []
    try:
        build_price_from_moysklad("test-token", tmp_path / "price.xlsx", client=client)
    except MoySkladError as error:
        assert "Не найдены склады" in str(error)
    else:
        raise AssertionError("Missing stores must stop price generation")
