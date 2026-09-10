from moysklad_price import MoySkladClient


def test_customer_order_print_templates_uses_dedicated_collections():
    client = object.__new__(MoySkladClient)
    calls = []

    def rows(endpoint, params=None):
        calls.append(endpoint)
        return [{"name": endpoint.rsplit("/", 1)[-1]}]

    client._rows = rows
    templates = client.customer_order_print_templates()

    assert calls == [
        "entity/customerorder/metadata/embeddedtemplate",
        "entity/customerorder/metadata/customtemplate",
    ]
    assert [value["name"] for value in templates] == ["embeddedtemplate", "customtemplate"]
