from datetime import date, timedelta

import pytest

from crm_bot import client_text, parse_task_date
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
    assert db.complete_task(1, task.id)
    assert db.tasks(1) == []

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
