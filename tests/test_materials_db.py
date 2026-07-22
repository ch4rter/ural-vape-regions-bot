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


def test_consistent_backup(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    product = db.add_product("OGGO VLIQ")
    section = db.add_section(product.id, "Мокапы")
    db.add_material(section.id, "photo", file_id="photo-id")

    backup_path = tmp_path / "backup" / "materials.sqlite3"
    db.backup_to(backup_path)
    restored = MaterialsDB(backup_path)
    assert restored.list_products()[0].name == "OGGO VLIQ"
    assert restored.list_materials(section.id)[0].file_id == "photo-id"


def test_access_by_id_and_username_binding(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    by_id = db.add_access_user("123456789")
    by_username = db.add_access_user("@Sales_Manager")

    assert db.authorize_user(123456789, None) is True
    assert db.authorize_user(777777777, "sales_manager") is True
    bound = db.get_access_user(by_username.id)
    assert bound.telegram_id == 777777777
    assert db.authorize_user(888888888, "sales_manager") is False

    db.delete_access_user(by_id.id)
    assert db.authorize_user(123456789, None) is False


def test_roles_chat_registry_and_settings(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    user = db.add_access_user("123456789")
    assert db.user_role(123456789) == "user"
    db.set_access_role(user.id, "junior_admin")
    assert db.user_role(123456789) == "junior_admin"

    db.upsert_client_chat(-100123, "Клиентский чат", "supergroup", True)
    assert db.get_client_chat(-100123).title == "Клиентский чат"
    assert db.list_client_chats(active_only=True)[0].chat_id == -100123
    db.upsert_client_chat(-100123, "Новое название", "supergroup", False)
    assert db.list_client_chats(active_only=True) == []
    assert db.get_client_chat(-100123).is_active is False


def test_waitlist_is_scoped_to_manager_and_remembers_matches(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    first = db.add_wait_entry(-100123, "Vape Shop", 101, "Андрей", "XROS 0.6 2мл")
    second = db.add_wait_entry(-100456, "Другой клиент", 202, "Матвей", "OGGO VLIQ")

    assert db.list_wait_entries(manager_id=101) == [first]
    assert {entry.id for entry in db.list_wait_entries()} == {first.id, second.id}
    assert db.wait_match_seen(first.id, "xros 0 6 2мл") is False
    db.record_wait_match(first.id, "xros 0 6 2мл")
    assert db.wait_match_seen(first.id, "xros 0 6 2мл") is True
    assert db.close_wait_entry(first.id, manager_id=202) is False
    assert db.close_wait_entry(first.id, manager_id=101) is True
    assert db.list_wait_entries(manager_id=101) == []

    db.set_setting("service_chat_id", "-100123")
    assert db.get_setting("service_chat_id") == "-100123"


def test_existing_access_table_gets_role_migration(tmp_path):
    path = tmp_path / "old.sqlite3"
    connection = sqlite3.connect(path)
    connection.execute(
        """CREATE TABLE access_users (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               telegram_id INTEGER UNIQUE,
               username TEXT UNIQUE,
               created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
               CHECK(telegram_id IS NOT NULL OR username IS NOT NULL)
           )"""
    )
    connection.execute("INSERT INTO access_users(telegram_id) VALUES (123456789)")
    connection.commit()
    connection.close()

    db = MaterialsDB(path)
    assert db.user_role(123456789) == "user"
