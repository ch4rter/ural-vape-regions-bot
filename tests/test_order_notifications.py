import asyncio
from decimal import Decimal

from bot import deliver_order_notice
from materials_db import MaterialsDB
from order_notifications import (
    checked_notification_targets,
    parse_order_notice,
    review_notification_targets,
)


def configured_database(tmp_path):
    db = MaterialsDB(tmp_path / "materials.sqlite3")
    manager = db.add_access_user("100000011")
    assistant = db.add_access_user("100000012")
    reviewer = db.add_access_user("100000013")
    db.set_access_role(manager.id, "manager")
    db.set_access_role(assistant.id, "assistant")
    db.set_access_role(reviewer.id, "junior_admin")
    channel = "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/valera"
    db.set_access_sales_channel(manager.id, "Валера", channel)
    db.set_team_assistant(manager.id, assistant.id)
    db.set_setting("order_review_reviewer_access_id", str(reviewer.id))
    return db, manager, assistant, reviewer, channel


def test_notification_routes_assistant_then_manager_then_admin(tmp_path):
    db, manager, assistant, reviewer, channel = configured_database(tmp_path)
    assert checked_notification_targets(db, channel, {999}) == (
        assistant.telegram_id, manager.telegram_id, 999,
    )
    assert review_notification_targets(db, {999}) == (reviewer.telegram_id, 999)
    assert checked_notification_targets(db, "", {999}) == (999,)


def test_order_notice_parsing_and_watch_state(tmp_path):
    notice = parse_order_notice({
        "id": "order-1",
        "name": "00001",
        "updated": "2026-08-22 12:00:00.000",
        "moment": "2026-08-22 11:00:00.000",
        "sum": 1234500,
        "agent": {"name": "ТРИНИТИ"},
        "state": {
            "name": "Проверено",
            "meta": {"href": "https://api.moysklad.ru/api/remap/1.2/entity/customerorder/metadata/states/checked"},
        },
        "salesChannel": {
            "name": "Валера",
            "meta": {"href": "https://api.moysklad.ru/api/remap/1.2/entity/saleschannel/valera"},
        },
        "meta": {"uuidHref": "https://online.moysklad.ru/app/#customerorder/edit?id=order-1"},
    })
    assert notice is not None
    assert notice.total == Decimal("12345")
    assert notice.agent == "ТРИНИТИ"
    assert notice.web_url.startswith("https://online.moysklad.ru/")

    db = MaterialsDB(tmp_path / "watch.sqlite3")
    db.save_order_watch_state(
        notice.order_id, notice.state_href, notice.state_name, notice.updated,
        notified=False,
    )
    assert db.order_watch_state("order-1")["notified_state_href"] is None
    db.save_order_watch_state(
        notice.order_id, notice.state_href, notice.state_name, notice.updated,
        notified=True,
    )
    assert db.order_watch_state("order-1")["notified_state_href"] == notice.state_href


def test_delivery_falls_back_from_assistant_to_manager():
    notice = parse_order_notice({
        "id": "order-2",
        "name": "00002",
        "sum": 10000,
        "state": {
            "name": "Проверено",
            "meta": {"href": "https://api.moysklad.ru/state/checked"},
        },
    })

    class FakeBot:
        def __init__(self):
            self.calls = []

        async def send_message(self, recipient, *args, **kwargs):
            self.calls.append(recipient)
            if recipient == 101:
                raise RuntimeError("assistant blocked bot")

    fake = FakeBot()
    delivered = asyncio.run(
        deliver_order_notice(fake, notice, (101, 202, 303), checked=True)
    )
    assert delivered == 202
    assert fake.calls == [101, 202]
