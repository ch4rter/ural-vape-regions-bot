from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from customer_activity import (
    ActivityDatabase,
    build_activity_excel,
    client_parts,
    report_rows,
    sync_shipments,
)


def demand(identifier, name, moment, channel="Валера", channel_id="v", amount=10000):
    return {
        "id": identifier, "name": identifier, "moment": moment,
        "updated": moment, "sum": amount,
        "agent": {"name": name, "meta": {"href": f"https://api.moysklad.ru/entity/counterparty/{identifier}"}},
        "salesChannel": {"name": channel, "meta": {"href": f"https://api.moysklad.ru/entity/saleschannel/{channel_id}"}},
    }


class FakeClient:
    def __init__(self, rows):
        self.rows = rows
        self.params = []

    def _rows(self, endpoint, params):
        self.params.append((endpoint, params))
        return self.rows


def test_client_names_are_grouped_by_prefix():
    assert client_parts("Is_getto/Шаров А.И.-5/Вологда")[0] == client_parts("Is getto/Литвиневский В.М.-5")[0]
    assert client_parts("(О) Vape Zone/НАЛ/СПб")[1] == "(О) Vape Zone"
    assert client_parts("Вячеслав   KAILAS {Ульяновск} | Нал")[0] == client_parts(
        "Вячеслав KAILAS {Ульяновск} | (ИП Ниськов А.Н.)"
    )[0]
    assert client_parts("Manager   Sistem Opt (Владимирская область)| (ИП Механиков Д. А.)")[0] == client_parts(
        "Manager Sistem Opt (Владимирская область) | (ИП Куликова А.С.)"
    )[0]
    assert client_parts("Joki   Липецк (ИП Нечаев В. А.)")[0] == client_parts(
        "Joki Липецк (ИП Лаврентьев   А.Н.)"
    )[0]
    assert client_parts("Joki Липецк (ИП Толчеев С.Г.)")[1] == "Joki Липецк"
    assert client_parts("Oleg Wow   Smoke {Самара} (ИП Исмаилов Р.Р.)")[0] == client_parts(
        "Oleg Wow Smoke {Самара} (ИП   Севостьянов Р.С.)"
    )[0]
    assert client_parts("Oleg Wow Smoke {Самара} (ИП Севостьянов Р.С.)")[1] == (
        "Oleg Wow Smoke {Самара}"
    )


def test_weekly_groups_and_latest_channel(tmp_path):
    db = ActivityDatabase(tmp_path / "activity.sqlite3")
    db.upsert_shipments([
        demand("old-return", "Return/ИП 1", "2026-06-01 10:00:00", "Андрей", "a"),
        demand("returned", "Return/ИП 2", "2026-09-03 10:00:00", "Валера", "v"),
        demand("new", "New shop/ИП", "2026-09-02 10:00:00", "Матвей", "m", 25000),
        demand("lost", "Lost/ИП", "2026-08-10 10:00:00", "Андрей", "a"),
    ])
    data = report_rows(db.shipments(), date(2026, 9, 7))
    assert [row["client"] for row in data["new"]] == ["New shop"]
    assert data["returned"][0]["client"] == "Return"
    assert data["returned"][0]["manager"] == "Валера"
    assert data["returned"][0]["gap"] == 94
    assert data["lost"][0]["category"] == "20–29 дней"
    assert len(data["previous"]["new"]) == 0
    assert len(data["previous"]["lost"]) == 2


def test_sync_uses_start_date_then_incremental_cursor(tmp_path):
    db = ActivityDatabase(tmp_path / "activity.sqlite3")
    client = FakeClient([demand("one", "Shop/ИП", "2026-08-01 10:00:00")])
    sync_shipments(db, client, datetime(2026, 9, 1, 3, 0))
    sync_shipments(db, client, datetime(2026, 9, 8, 3, 0))
    assert "moment>=2026-04-01 00:00:00" in client.params[0][1]["filter"]
    assert client.params[0][1]["limit"] == 100
    assert "updated>=" not in client.params[0][1]["filter"]
    assert "updated>=2026-09-01 02:59:55" in client.params[1][1]["filter"]


def test_excel_has_dashboard_and_single_lost_sheet(tmp_path):
    destination = tmp_path / "report.xlsx"
    db = ActivityDatabase(tmp_path / "excel.sqlite3")
    db.upsert_shipments([demand("lost", "Lost/ИП", "2026-08-01 10:00:00")])
    data = report_rows(db.shipments(), date(2026, 9, 7))
    build_activity_excel(destination, data)
    workbook = load_workbook(destination)
    assert workbook.sheetnames == ["Дэшборд", "Помесячная аналитика", "Новые кенты", "Вернувшиеся кенты", "Кенты-потеряшки"]
    dashboard = workbook["Дэшборд"]
    assert len(dashboard._charts) == 2
    assert len(dashboard._charts[0].ser) == 2
    assert len(dashboard._charts[1].ser) == 1
    assert dashboard["A25"].value == "Как читать дэшборд"
    monthly_headers = [cell.value for cell in workbook["Помесячная аналитика"][1]]
    assert "Уникальные покупатели" in monthly_headers
    assert "Выручка на покупателя" in monthly_headers
    assert workbook["Помесячная аналитика"].max_row == 6  # April through August only.
    headers = [cell.value for cell in dashboard[20]]
    assert dashboard["F19"].value == "Клиенты, которые давно не заказывали"
    assert "Новые — предыдущая" not in headers
    assert "Вернувшиеся — предыдущая" not in headers
    assert "Активные неделю назад" not in headers
    assert dashboard.freeze_panes == "A3"
