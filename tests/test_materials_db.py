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


def test_material_album_is_stored_ordered_and_deleted_as_one_item(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    product = db.add_product("OGGO")
    section = db.add_section(product.id, "Мокапы")
    second = db.add_material(
        section.id, "photo", file_id="photo-2", media_group_id="chat:album",
        media_group_position=102,
    )
    first = db.add_material(
        section.id, "photo", file_id="photo-1", caption="Мокапы",
        media_group_id="chat:album", media_group_position=101,
    )

    stored = db.list_materials(section.id)
    assert {item.media_group_id for item in stored} == {"chat:album"}
    assert {item.media_group_position for item in stored} == {101, 102}

    db.delete_material(second.id)
    assert db.get_material(first.id) is None
    assert db.list_materials(section.id) == []


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
    assert db.user_role(123456789) == "employee"
    db.set_access_role(user.id, "junior_admin")
    assert db.user_role(123456789) == "junior_admin"
    channel_href = "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/valera"
    db.set_access_sales_channel(user.id, "Валера", channel_href)
    assert db.user_sales_channel(123456789) == ("Валера", channel_href)
    assert db.get_access_user(user.id).sales_channel_name == "Валера"

    db.upsert_client_chat(-100123, "Клиентский чат", "supergroup", True)
    assert db.get_client_chat(-100123).title == "Клиентский чат"
    assert db.list_client_chats(active_only=True)[0].chat_id == -100123
    db.upsert_client_chat(-100123, "Новое название", "supergroup", False)
    assert db.list_client_chats(active_only=True) == []
    assert db.get_client_chat(-100123).is_active is False


def test_client_chat_tags_use_or_filter_and_keep_actor(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    db.upsert_client_chat(
        -100101, "БП", "supergroup", True, added_by_id=11, added_by_username="manager"
    )
    db.upsert_client_chat(-100202, "Железо", "supergroup", True)
    db.upsert_client_chat(-100303, "Оба", "supergroup", True)
    db.upsert_client_chat(-100404, "Без тегов", "supergroup", True)
    db.set_client_chat_tags(-100101, ["bp"])
    db.set_client_chat_tags(-100202, ["hardware"])
    db.set_client_chat_tags(-100303, ["bp", "hardware"])

    selected = db.list_client_chats(active_only=True, tags=["bp", "hardware"])
    assert {chat.chat_id for chat in selected} == {-100101, -100202, -100303}
    assert db.get_client_chat(-100101).added_by_id == 11
    assert db.get_client_chat(-100101).added_by_username == "manager"
    assert db.get_client_chat(-100303).tags == ("bp", "hardware")
    assert [chat.chat_id for chat in db.list_client_chats(untagged_only=True)] == [-100404]

    assert db.toggle_client_chat_tag(-100202, "hardware") == ()
    assert db.toggle_client_chat_tag(-100202, "sp") == ("sp",)


def test_public_sections_are_opt_in(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    product = db.add_product("OGGO VLIQ")
    declarations = db.add_section(product.id, "Декларации")
    internal = db.add_section(product.id, "Внутреннее КП")
    assert declarations.is_public is False
    assert db.list_public_products() == []
    assert db.list_sections(product.id, public_only=True) == []

    assert db.toggle_section_public(declarations.id) is True
    assert db.list_public_products() == [product]
    assert db.list_sections(product.id, public_only=True)[0].name == "Декларации"
    assert db.get_section(internal.id).is_public is False
    assert db.toggle_section_public(declarations.id) is False
    assert db.list_public_products() == []


def test_lead_profile_is_saved_once(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    profile = db.save_lead_profile(
        123456, "client_user", "Иван", "Владелец", "Тамбов",
        "Vape Shop", "3 точки + опт", "Андрей",
    )
    assert profile.manager == "Андрей"
    assert db.get_lead_profile(123456).company == "Vape Shop"

    repeated = db.save_lead_profile(
        123456, "changed", "Другое имя", "Закупщик", "Москва",
        "Другая компания", "опт", "Матвей",
    )
    assert repeated.full_name == "Иван"
    assert len(db.list_lead_profiles()) == 1


def test_unmatched_regions_are_stored_separately_and_resolved_together(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    first = db.save_unmatched_region(101, "client_one", "  Новый город ", "новый город")
    repeated = db.save_unmatched_region(101, "client_one", "Новый город", "новый город")
    second = db.save_unmatched_region(202, None, "НОВЫЙ ГОРОД", "новый город")

    assert first.id == repeated.id
    assert len(db.list_unmatched_regions()) == 2
    assert db.get_unmatched_region(second.id).raw_region == "НОВЫЙ ГОРОД"
    assert db.resolve_unmatched_regions("новый город") == 2
    assert db.list_unmatched_regions() == []


def test_known_telegram_user_resolves_manager_username(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    db.remember_telegram_user(101, "ShmidtUV", "Андрей")
    assert db.telegram_user_id_by_username("@shmidtuv") == 101

    db.remember_telegram_user(101, None, "Андрей Шмидт")
    assert db.telegram_user_id_by_username("SHMIDTUV") == 101


def test_private_bot_users_are_tracked_separately(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    db.remember_telegram_user(101, "client", "Клиент")
    assert db.list_activated_telegram_users() == []
    assert db.list_unconfirmed_telegram_users() == [(101, "client", "Клиент")]
    db.mark_telegram_user_activated(101)
    users = db.list_activated_telegram_users()
    assert len(users) == 1
    assert users[0].user_id == 101
    assert users[0].full_name == "Клиент"
    assert db.list_unconfirmed_telegram_users() == []


def test_waitlist_is_scoped_to_manager_and_remembers_matches(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    first = db.add_wait_entry(
        -100123, "Vape Shop", 101, "Андрей", "XROS 0.6 2мл", 777, "Нужно 10 упаковок"
    )
    second = db.add_wait_entry(-100456, "Другой клиент", 202, "Матвей", "OGGO VLIQ")
    manual_one = db.add_wait_entry(0, "Клиент без чата", 101, "Андрей", "OGGO VLIQ")
    manual_two = db.add_wait_entry(0, "Ещё один клиент", 101, "Андрей", "OGGO VLIQ")

    assert {entry.id for entry in db.list_wait_entries(manager_id=101)} == {
        first.id, manual_one.id, manual_two.id,
    }
    assert first.comment == "Нужно 10 упаковок"
    assert first.source_message_id == 777
    duplicate = db.add_wait_entry(-100123, "Vape Shop", 101, "Андрей", "xros 0.6 2МЛ")
    assert duplicate.id == first.id
    assert len(db.list_wait_entries(manager_id=101)) == 3
    assert {entry.id for entry in db.list_wait_entries()} == {
        first.id, second.id, manual_one.id, manual_two.id,
    }
    assert db.wait_match_seen(first.id, "xros 0 6 2мл") is False
    db.record_wait_match(first.id, "xros 0 6 2мл")
    assert db.wait_match_seen(first.id, "xros 0 6 2мл") is True
    db.set_wait_last_match(first.id, {"count": 2, "warehouses": ["Москва"], "items": []})
    assert db.get_wait_entry(first.id).last_match["count"] == 2
    updated = db.update_wait_entry(first.id, 101, query="картриджи XROS", comment="Любые варианты")
    assert updated.query == "картриджи XROS"
    assert updated.comment == "Любые варианты"
    assert updated.last_match is None
    assert db.wait_match_seen(first.id, "xros 0 6 2мл") is False
    assert db.close_wait_entry(first.id, manager_id=202) is False
    assert db.close_wait_entry(first.id, manager_id=101) is True
    assert {entry.id for entry in db.list_wait_entries(manager_id=101)} == {
        manual_one.id, manual_two.id,
    }

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
    assert db.user_role(123456789) == "employee"
    assert db.user_sales_channel(123456789) is None


def test_employee_permissions_and_sales_teams(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    manager_one = db.add_access_user("100000001")
    manager_two = db.add_access_user("100000002")
    assistant = db.add_access_user("100000003")
    db.set_access_role(manager_one.id, "manager")
    db.set_access_role(manager_two.id, "manager")
    db.set_access_role(assistant.id, "assistant")
    db.set_access_sales_channel(manager_one.id, "Валера", "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/1")
    db.set_access_sales_channel(manager_two.id, "Андрей", "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/2")
    db.set_team_assistant(manager_one.id, assistant.id)
    db.set_team_assistant(manager_two.id, assistant.id)

    assert db.team_assistant(manager_one.id).telegram_id == 100000003
    assert db.team_assistant(manager_two.id).telegram_id == 100000003
    assert len(db.list_sales_teams()) == 2
    assert db.access_user_by_channel(
        "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/1?expand=x"
    ).id == manager_one.id

    db.set_access_permission(assistant.id, "manage_prices", True)
    assert db.user_permissions(100000003) == {"manage_prices"}
    db.set_access_permission(assistant.id, "manage_prices", False)
    assert db.user_permissions(100000003) == set()


def test_legacy_junior_permissions_are_seeded_only_once(tmp_path):
    path = tmp_path / "legacy.sqlite3"
    db = MaterialsDB(path)
    user = db.add_access_user("100000004")
    db.set_access_role(user.id, "junior_admin")
    # The migration already ran before this user was promoted, so permissions
    # remain explicitly controlled rather than being silently restored.
    db.set_access_permission(user.id, "broadcasts", True)
    db.set_access_permission(user.id, "broadcasts", False)
    reopened = MaterialsDB(path)
    assert reopened.permissions_for_access(user.id) == set()
