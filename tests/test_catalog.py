from pathlib import Path

from openpyxl import Workbook

from bot import Catalog, Entry, format_result, normalize, suggestion_label


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


def test_new_three_column_format(tmp_path):
    path = tmp_path / "managers.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Местоположение", "Территория", "Менеджер"])
    sheet.append(["Россия / Костромская область", "Питер", "Андрей"])
    sheet.append(["Россия / Нижегородская область", "Питер", "Елена"])
    workbook.save(path)

    catalog = Catalog(path)
    found = catalog.exact("питер")
    assert len(found) == 2
    assert found[0].location == "Россия / Костромская область"
    result = format_result(found)
    assert "Местоположение" in result
    assert "Нижегородская область" in result
    assert suggestion_label(found[0]) == "Питер — Россия / Костромская область"


def test_many_results_are_paginated():
    entries = [Entry("Александровка", "Андрей", f"Район {i}") for i in range(25)]
    first_page = format_result(entries)
    last_page = format_result(entries, page=2)
    assert "Страница <b>1</b> из <b>3</b>" in first_page
    assert "Район 9" in first_page
    assert "Район 10" not in first_page
    assert "Район 24" in last_page
    assert suggestion_label(entries[0], 25) == "Александровка — 25 вариантов"
