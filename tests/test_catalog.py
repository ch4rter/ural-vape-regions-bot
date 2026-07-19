from pathlib import Path

from openpyxl import Workbook

from bot import Catalog, normalize


def make_book(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Название", "Менеджер"])
    sheet.append(["Орёл", "Анна"])
    sheet.append(["  Нижний   Новгород ", "Иван"])
    sheet.append(["Орёл", "Пётр"])
    workbook.save(path)


def test_normalize():
    assert normalize("  ОРЁЛ — город ") == "орел город"


def test_exact_and_duplicates(tmp_path):
    path = tmp_path / "managers.xlsx"
    make_book(path)
    catalog = Catalog(path)
    assert [item.manager for item in catalog.exact("орел")] == ["Анна", "Пётр"]
    assert catalog.exact("нижний новгород")[0].manager == "Иван"


def test_suggestion(tmp_path):
    path = tmp_path / "managers.xlsx"
    make_book(path)
    catalog = Catalog(path)
    indexes = catalog.suggestions("Нижний Новгорд")
    assert catalog.entries[indexes[0]].manager == "Иван"
