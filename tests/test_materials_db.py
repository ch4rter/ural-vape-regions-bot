import sqlite3

from materials_db import MaterialsDB


def test_product_section_material_lifecycle(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    product = db.add_product("OGGO VLIQ")
    section = db.add_section(product.id, "Декларации")
    text = db.add_material(section.id, "text", text="Коммерческое предложение")
    document = db.add_material(
        section.id,
        "document",
        file_id="telegram-file-id",
        caption="Декларация",
        file_name="declaration.pdf",
    )

    assert db.list_products(visible_only=True) == [product]
    assert db.list_sections(product.id) == [section]
    assert [item.id for item in db.list_materials(section.id)] == [text.id, document.id]

    db.delete_product(product.id)
    assert db.list_products() == []
    assert db.get_section(section.id) is None
    assert db.get_material(document.id) is None


def test_hidden_products_and_unique_names(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    product = db.add_product("OGGO VLIQ")
    db.toggle_product(product.id)
    assert db.list_products(visible_only=True) == []
    assert db.list_products()[0].is_visible is False

    try:
        db.add_product("oggo vliq")
    except sqlite3.IntegrityError:
        pass
    else:
        raise AssertionError("Product names must be unique ignoring case")
