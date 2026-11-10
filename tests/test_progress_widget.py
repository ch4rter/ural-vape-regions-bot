import asyncio
import json
import zlib
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock

import pytest
from aiogram import Bot, Dispatcher
from aiogram.exceptions import TelegramBadRequest
from aiogram.types import Update

import bot as app
from progress_widget import (
    AnswerWidgetQuery, EditMessageWidgetData, ProgressWidgets, SendWidget,
    WidgetQuery, build_state, encode_state,
)
from sales_progress import ManagerProgress


def make_service():
    settings = {}
    service = ProgressWidgets(
        access=lambda uid, username: uid == 123,
        months=lambda: ["2026-10", "2026-11"],
        snapshot=AsyncMock(return_value={"m": "2026-11", "rows": []}),
        get_setting=settings.get, set_setting=settings.__setitem__,
    )
    settings[service.key(123, 4)] = "1"
    return service


def query(**changes):
    raw = dict(id="q1", **{"from": {"id": 123, "first_name": "Test", "is_bot": False}},
               chat={"id": 123, "type": "private"}, message_id=4,
               data="pw:2026-11:0:refresh", type={"type": "data"})
    raw.update(changes)
    return WidgetQuery.model_validate(raw)


def test_widget_sizes_financial_totals_and_all_managers():
    html = Path("progress_widget.html").read_bytes()
    assert len(html) <= 32768
    assert len(zlib.compress(html, 9)) <= 4096
    rows = [ManagerProgress("Менеджер " + str(i), str(i), Decimal(i * 100)) for i in range(8)]
    state = build_state("2026-11", ["2026-10", "2026-11"], "Ноябрь 2026",
                        Decimal(2000), rows, now=datetime(2026, 11, 10))
    assert state["total"] == "2 800 ₽"
    assert state["pct"] == "140.0"
    assert state["left"] == "0 ₽"
    assert state["rows"][0] == ["Менеджер 7", "700 ₽", "25.0"]
    assert len(state["rows"]) == 8
    assert "pages" not in state and "updated" not in state and "timing" not in state
    assert state["delta"] == "+2 200 ₽"
    assert state["prev"] == "2026-10" and state["next"] == ""
    assert len(encode_state(state).encode()) <= 1024
    state = build_state("2026-11", ["2026-11"], "Ноябрь", None, [], now=datetime(2026, 11, 10))
    assert "error" in state and "total" not in state


@pytest.mark.parametrize("changes", [
    {"from": {"id": 456, "first_name": "Other", "is_bot": False}},
    {"chat": {"id": -100, "type": "supergroup"}},
    {"message_id": 99},
    {"data": "pw:1900-01:0:refresh"},
    {"data": "pw:2026-11:-1:show"},
    {"data": "pw:2026-11:0:delete"},
    {"type": {"type": "sent_transaction"}},
])
def test_rejects_unauthorized_context_and_invalid_queries(changes):
    service, telegram = make_service(), AsyncMock()
    asyncio.run(service.handle(query(**changes), telegram))
    service.snapshot.assert_not_called()
    assert telegram.call_count == 1
    assert isinstance(telegram.call_args.args[0], AnswerWidgetQuery)


def test_ack_precedes_loading_and_only_private_state_is_edited():
    service, telegram = make_service(), AsyncMock()

    async def snapshot(*args):
        assert isinstance(telegram.call_args.args[0], AnswerWidgetQuery)
        return {"m": "2026-11", "total": "100 ₽"}

    service.snapshot.side_effect = snapshot
    asyncio.run(service.handle(query(), telegram))
    edit = telegram.call_args.args[0]
    assert isinstance(edit, EditMessageWidgetData)
    assert edit.chat_id == 123 and edit.message_id == 4
    assert "public_data" not in edit.model_dump()
    assert json.loads(edit.chat_data)["total"] == "100 ₽"
    assert not service.busy


def test_load_error_is_shown_without_exception_details():
    service, telegram = make_service(), AsyncMock()
    service.snapshot.side_effect = RuntimeError("secret URL")
    asyncio.run(service.handle(query(), telegram))
    state = json.loads(telegram.call_args.args[0].chat_data)
    assert "Не удалось" in state["error"]
    assert "secret" not in state["error"]


def test_revoked_access_during_load_clears_display():
    service, telegram = make_service(), AsyncMock()

    async def snapshot(*args):
        service.access = lambda *args: False
        return {"total": "100 ₽"}

    service.snapshot.side_effect = snapshot
    asyncio.run(service.handle(query(), telegram))
    state = json.loads(telegram.call_args.args[0].chat_data)
    assert "total" not in state
    assert "Доступ закрыт" in state["error"]


def test_aiogram_dispatcher_receives_unknown_widget_update():
    async def run():
        service, telegram = make_service(), Bot("123:TEST")
        dispatcher = Dispatcher()
        dispatcher.update.outer_middleware(service)
        service.handle = AsyncMock()
        update = Update.model_validate({"update_id": 1, "widget_query": query().model_dump(by_alias=True)},
                                       context={"bot": telegram})
        await dispatcher.feed_update(telegram, update)
        service.handle.assert_awaited_once()
        assert service.handle.call_args.args[0].from_user.id == 123
        await telegram.session.close()
        await dispatcher.storage.close()
    asyncio.run(run())


def test_send_falls_back_when_api_does_not_support_widgets():
    service = make_service()
    telegram = AsyncMock(side_effect=TelegramBadRequest(method=SimpleNamespace(), message="unsupported"))
    message = SimpleNamespace(chat=SimpleNamespace(id=123, type="private"),
                              from_user=SimpleNamespace(id=123, username=None),
                              message_thread_id=None, bot=telegram)
    assert asyncio.run(service.send(message)) is False
    assert isinstance(telegram.call_args.args[0], SendWidget)


def test_parallel_refresh_reuses_single_moysklad_load(monkeypatch):
    async def run():
        monkeypatch.setattr(app, "sales_progress_cache", {})
        monkeypatch.setattr(app, "sales_progress_locks", {})
        fetch = AsyncMock(return_value=[])

        async def load(month, *, force=False):
            if month in app.sales_progress_cache and not force:
                return app.sales_progress_cache[month][1]
            await fetch()
            await asyncio.sleep(0.01)
            app.sales_progress_cache[month] = (datetime.now(), [])
            return []

        monkeypatch.setattr(app, "_load_progress_rows", load)
        await asyncio.gather(*(app.load_progress_rows("2026-11", force=True) for _ in range(5)))
        fetch.assert_awaited_once()
    asyncio.run(run())


def test_missing_goal_does_not_query_moysklad(monkeypatch):
    monkeypatch.setattr(app, "materials_db", SimpleNamespace(get_setting=lambda key: None), raising=False)
    monkeypatch.setattr(app, "progress_months", lambda: ["2026-11"])
    load = AsyncMock()
    monkeypatch.setattr(app, "load_progress_rows", load)
    state = asyncio.run(app.progress_widget_snapshot("2026-11", 0, False))
    assert "error" in state
    load.assert_not_called()


@pytest.mark.parametrize("chat_type,access,widget_sent,expected_text", [
    ("private", True, True, True),
    ("private", True, False, True),
    ("supergroup", True, True, True),
    ("private", False, True, False),
])
def test_progress_command_routing(monkeypatch, chat_type, access, widget_sent, expected_text):
    monkeypatch.setattr(app, "has_internal_access", lambda *args: access)
    send = AsyncMock(return_value=widget_sent)
    monkeypatch.setattr(app.progress_widgets, "send", send)
    message = SimpleNamespace(chat=SimpleNamespace(type=chat_type),
                              from_user=SimpleNamespace(id=123, username=None), answer=AsyncMock())
    asyncio.run(app.command_sales_progress(message))
    assert message.answer.called == expected_text
    send.assert_not_called()


@pytest.mark.parametrize("chat_type,access,sent,text", [
    ("private", True, True, False), ("private", True, False, True),
    ("supergroup", True, True, True), ("private", False, True, False),
])
def test_progress2_separate_command(monkeypatch, chat_type, access, sent, text):
    monkeypatch.setattr(app, "has_internal_access", lambda *args: access)
    send = AsyncMock(return_value=sent)
    monkeypatch.setattr(app.progress_widgets, "send", send)
    message = SimpleNamespace(chat=SimpleNamespace(type=chat_type),
                              from_user=SimpleNamespace(id=123, username=None), answer=AsyncMock())
    asyncio.run(app.command_sales_progress_widget(message))
    assert message.answer.called == text
    assert send.called == (chat_type == "private" and access)


def test_widget_has_no_caption_or_external_buttons():
    service = make_service()
    telegram = AsyncMock(return_value=SimpleNamespace(chat=SimpleNamespace(id=123), message_id=5))
    message = SimpleNamespace(chat=SimpleNamespace(id=123, type="private"),
                              from_user=SimpleNamespace(id=123, username=None),
                              message_thread_id=None, bot=telegram)
    assert asyncio.run(service.send(message)) is True
    payload = telegram.call_args.args[0].model_dump(exclude_none=True)
    assert "caption" not in payload and "reply_markup" not in payload


@pytest.mark.parametrize("month,now,phase,delta", [
    ("2026-11", datetime(2026, 11, 1), "current", "+0 ₽"),
    ("2026-11", datetime(2026, 11, 16), "current", "−1 500 ₽"),
    ("2026-10", datetime(2026, 11, 16), "past", "−3 000 ₽"),
    ("2026-12", datetime(2026, 11, 16), "future", "+0 ₽"),
])
def test_calendar_benchmark_and_zero_sales(month, now, phase, delta):
    state = build_state(month, [month], month, Decimal(3000),
                        [ManagerProgress("Тест", "x", Decimal(0))], now=now)
    assert state["phase"] == phase and state["delta"] == delta
    assert state["rows"][0][2] == "0"
