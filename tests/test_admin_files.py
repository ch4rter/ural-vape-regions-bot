import zipfile

import bot
from materials_db import MaterialsDB
from openpyxl import Workbook


def make_excel(path, territory):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "База бота"
    sheet.append(["Местоположение", "Территория", "Менеджер"])
    sheet.append(["Россия / Тестовая область", territory, "Андрей"])
    workbook.save(path)


def test_backup_archive_contains_database_and_active_excel(tmp_path, monkeypatch):
    database = MaterialsDB(tmp_path / "data" / "materials.sqlite3")
    database.add_product("OGGO VLIQ")
    excel = tmp_path / "managers.xlsx"
    make_excel(excel, "Питер")

    monkeypatch.setattr(bot, "materials_db", database, raising=False)
    monkeypatch.setattr(bot, "active_excel_path", excel, raising=False)
    monkeypatch.setattr(bot, "catalog", bot.Catalog(excel), raising=False)
    archive = tmp_path / "backup.zip"
    bot.build_backup_archive(archive)

    with zipfile.ZipFile(archive) as backup:
        assert set(backup.namelist()) == {"materials.sqlite3", "managers.xlsx", "README.txt"}


def test_applying_excel_keeps_database_and_backs_up_old_excel(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    old_excel = tmp_path / "seed.xlsx"
    pending_excel = tmp_path / "pending.xlsx"
    managed_excel = data_dir / "managers.xlsx"
    make_excel(old_excel, "Старый город")
    make_excel(pending_excel, "Новый город")
    database = MaterialsDB(data_dir / "materials.sqlite3")
    database.add_product("OGGO VLIQ")

    monkeypatch.setattr(bot, "active_excel_path", old_excel, raising=False)
    monkeypatch.setattr(bot, "managed_excel_path", managed_excel, raising=False)
    new_catalog, backup_path = bot.apply_pending_excel(pending_excel)

    assert new_catalog.exact("Новый город")
    assert backup_path.exists()
    assert bot.Catalog(backup_path).exact("Старый город")
    assert database.list_products()[0].name == "OGGO VLIQ"
