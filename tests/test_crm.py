import asyncio
from datetime import date, timedelta
from types import SimpleNamespace

import pytest

import crm_bot
from crm_bot import client_text, matches_filter, parse_task_date
from crm_db import CRMDatabase
from google_crm import CRMClient, GoogleCRM


def sample_client(**changes):
    values = {
        "row": 12,
        "name": "Vape House",
        "region": "Москва",
        "outlets": "3",
        "status": "Работаем",
        "client_type": "Розница",
        "contact_name": "Андрей",
        "phone": "79990000000",
        "telegram": "@vapehouse",
        "order_date": "01.08.2026",
        "call_date": "02.08.2026",
        "products": {"iron": "buy", "dojo": "no", "oggo": "unknown", "elfliq": "buy", "cosmo": "no"},
        "product_notes": {"iron": "", "dojo": "дорого", "oggo": "", "elfliq": "констр", "cosmo": "нет спроса"},
        "note": "Последний результат",
    }
    values.update(changes)
    return CRMClient(**values)


def test_crm_search_is_fuzzy_and_uses_contact_fields():
    service = GoogleCRM.__new__(GoogleCRM)
    clients = [sample_client(), sample_client(row=20, name="Другая компания", telegram="@other")]
    service.clients = lambda: clients

    assert service.search("vape hause")[0].name == "Vape House"
    assert service.search("vapehouse")[0].telegram == "@vapehouse"
    assert service.search("москва андрей")[0].name == "Vape House"


def test_crm_database_tasks_events_and_call_list(tmp_path):
    db = CRMDatabase(tmp_path / "crm.sqlite3")
    client = sample_client()
    task = db.add_task(1, client.identity, client.name, date.today(), "Позвонить")
    assert db.tasks(1)[0].text == "Позвонить"
    future = db.add_task(1, client.identity, client.name, date.today() + timedelta(days=10), "Будущее")
    assert [item.id for item in db.tasks(1)] == [task.id]
    assert [item.id for item in db.tasks(1, include_future=True)] == [task.id, future.id]
    assert db.complete_task(1, task.id)
    assert db.tasks(1) == []
    assert [item.id for item in db.tasks(1, include_future=True)] == [future.id]

    db.add_event(1, client.identity, client.name, "note", "Созвонились")
    assert db.daily_events(1)[0]["summary"] == "Созвонились"
    assert db.add_to_call_list(1, client.identity, client.name)
    assert not db.add_to_call_list(1, client.identity, client.name)
    assert db.call_list(1)[0]["client_name"] == client.name
    db.remove_from_call_list(1, client.identity)
    assert db.call_list(1) == []


def test_task_date_and_client_card():
    assert parse_task_date("сегодня") == date.today()
    assert parse_task_date("завтра") == date.today() + timedelta(days=1)
    assert parse_task_date("не дата") is None
    text = client_text(sample_client())
    assert "Vape House" in text
    assert "🟢 Железо" in text
    assert "🔴 DOJO" in text
    assert "⚪ OGGO" in text


def test_red_product_requires_reason():
    service = GoogleCRM.__new__(GoogleCRM)
    with pytest.raises(ValueError, match="причину"):
        service.set_product(12, sample_client().identity, "dojo", "no", "")


def test_configurable_client_filters():
    retail = sample_client(outlets="3", client_type="Розница")
    wholesale = sample_client(
        row=20, outlets="опт", client_type="Опт",
        products={"iron": "buy", "dojo": "buy", "oggo": "no", "elfliq": "no", "cosmo": "unknown"},
    )
    assert matches_filter(retail, {"kind": "retail", "scale": "1_3", "products": ["iron"], "product_mode": "all"})
    assert not matches_filter(retail, {"kind": "retail", "scale": "4_10", "products": [], "product_mode": "any"})
    assert matches_filter(wholesale, {"kind": "wholesale", "scale": "any", "products": ["dojo", "oggo"], "product_mode": "any"})
    assert not matches_filter(wholesale, {"kind": "wholesale", "scale": "any", "products": ["dojo", "oggo"], "product_mode": "all"})
    assert matches_filter(retail, {
        "kind": "all", "scale": "any", "region": "моск", "statuses": ["Работаем"],
        "products": ["dojo"], "product_mode": "all", "product_state": "no",
    })
    assert not matches_filter(retail, {
        "kind": "all", "scale": "any", "statuses": ["Ждем заказ"], "products": [],
    })
    assert matches_filter(sample_client(status=""), {
        "kind": "all", "scale": "any", "statuses": ["__empty__"], "products": [],
    })


def test_daily_report_groups_client_and_hides_technical_product_changes():
    events = [
        {"client_identity": "trinity", "client_name": "ТРИНИТИ", "kind": "note", "summary": "Созвонились, согласовали заказ"},
        {"client_identity": "trinity", "client_name": "ТРИНИТИ", "kind": "order", "summary": "Заказ сегодня"},
        {"client_identity": "trinity", "client_name": "ТРИНИТИ", "kind": "product", "summary": "DOJO: не покупает — дорого"},
        {"client_identity": "trinity", "client_name": "ТРИНИТИ", "kind": "product", "summary": "DOJO: покупает"},
    ]

    class FakeDatabase:
        def daily_events(self, owner_id):
            return events

    class FakeMessage:
        def __init__(self):
            self.texts = []

        async def edit_text(self, text, reply_markup=None):
            self.texts.append(text)

        async def answer(self, text, reply_markup=None):
            self.texts.append(text)

    class FakeCallback:
        def __init__(self):
            self.from_user = SimpleNamespace(id=5533726476)
            self.message = FakeMessage()

        async def answer(self, *args, **kwargs):
            pass

    class FakeState:
        async def clear(self):
            pass

    old_database, old_crm, old_owners = crm_bot.database, crm_bot.crm, crm_bot.owner_ids
    crm_bot.database, crm_bot.crm, crm_bot.owner_ids = FakeDatabase(), object(), {5533726476}
    callback = FakeCallback()
    try:
        asyncio.run(crm_bot.daily_report(callback, FakeState()))
    finally:
        crm_bot.database, crm_bot.crm, crm_bot.owner_ids = old_database, old_crm, old_owners

    report = "\n".join(callback.message.texts)
    assert report.count("ТРИНИТИ") == 1
    assert "Созвонились, согласовали заказ" in report
    assert "Сделал заказ" in report
    assert "не покупает" not in report
    assert "DOJO: покупает" not in report
