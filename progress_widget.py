"""Private /progress2 widget and a narrow Bot API 10.4 adapter for aiogram 3.29.

Protocol: pw:<YYYY-MM>:<page>:<show|refresh>. Financial data is chat_data only.
Docs: https://core.telegram.org/bots/widgets
"""
from __future__ import annotations

import json
import calendar
import logging
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from aiogram import BaseMiddleware
from aiogram.exceptions import TelegramBadRequest, TelegramNotFound
from aiogram.methods import TelegramMethod
from aiogram.types import Chat, Message, User
from pydantic import BaseModel, Field

from sales_progress import format_money


class SendWidget(TelegramMethod[Message]):
    __api_method__ = "sendWidget"
    __returning__ = Message
    chat_id: int
    html: str
    chat_data: str
    allow_queries: bool = True
    protect_content: bool = True
    message_thread_id: int | None = None


class EditMessageWidgetData(TelegramMethod[bool]):
    __api_method__ = "editMessageWidgetData"
    __returning__ = bool
    chat_id: int
    message_id: int
    chat_data: str


class AnswerWidgetQuery(TelegramMethod[bool]):
    __api_method__ = "answerWidgetQuery"
    __returning__ = bool
    widget_query_id: str
    data: str


class WidgetQuery(BaseModel):
    id: str
    from_user: User = Field(alias="from")
    chat: Chat | None = None
    message_id: int | None = None
    data: str
    type: dict[str, Any]


def encode_state(state: dict) -> str:
    result = json.dumps(state, ensure_ascii=False, separators=(",", ":"))
    if len(result.encode("utf-8")) > 1024:
        raise ValueError("Progress widget state exceeds 1024 bytes")
    return result


def build_state(month, months, title, goal, rows, *, now):
    rows = sorted(rows, key=lambda row: (-row.revenue, row.name.casefold()))
    index = months.index(month)
    state = dict(m=month, title=title, prev=months[index - 1] if index else "",
                 next=months[index + 1] if index + 1 < len(months) else "",
                 rows=[])
    if goal is None:
        state["error"] = "Цель на этот месяц ещё не установлена."
        return state
    total = sum((row.revenue for row in rows), Decimal(0))
    remaining = max(goal - total, Decimal(0))
    year, number = map(int, month.split("-"))
    days = calendar.monthrange(year, number)[1]
    # Linear calendar benchmark at the start of today, not a sales forecast.
    elapsed = 0 if month > now.strftime("%Y-%m") else (
        days if month < now.strftime("%Y-%m") else now.day - 1)
    pace = goal * Decimal(elapsed) / Decimal(days)
    delta = total - pace
    state.update(pct=str((total / goal * 100).quantize(Decimal("0.1"))),
                 total=format_money(total), goal=format_money(goal),
                 left=format_money(remaining),
                 delta=("+" if delta >= 0 else "−") + format_money(abs(delta)),
                 phase="future" if month > now.strftime("%Y-%m") else ("past" if elapsed == days else "current"),
                 rows=[[row.name[:32], format_money(row.revenue),
                        str((row.revenue / total * 100).quantize(Decimal("0.1"))) if total else "0"]
                       for row in rows])
    return state


class ProgressWidgets(BaseMiddleware):
    def __init__(self, *, access, months, snapshot, get_setting, set_setting):
        self.access, self.months, self.snapshot = access, months, snapshot
        self.get_setting, self.set_setting = get_setting, set_setting
        self.busy: set[int] = set()
        self.html = Path(__file__).with_name("progress_widget.html").read_text(encoding="utf-8")

    @staticmethod
    def key(chat_id, message_id):
        return f"progress_widget:v2:{chat_id}:{message_id}"

    async def send(self, message):
        user = message.from_user
        if message.chat.type != "private" or not user or not self.access(user.id, user.username):
            return False
        month = datetime.now(ZoneInfo("Europe/Moscow")).strftime("%Y-%m")
        state = await self.snapshot(month, 0, False)
        if not self.access(user.id, user.username):
            return True
        try:
            sent = await message.bot(SendWidget(
                chat_id=message.chat.id, html=self.html, chat_data=encode_state(state),
                message_thread_id=message.message_thread_id,
            ))
        except (TelegramBadRequest, TelegramNotFound) as error:
            logging.warning("Progress widget unavailable (%s); using text", type(error).__name__)
            return False
        self.set_setting(self.key(sent.chat.id, sent.message_id), "1")
        return True

    async def __call__(self, handler, event, data):
        raw = getattr(event, "widget_query", None)
        if raw is None:
            return await handler(event, data)
        query = WidgetQuery.model_validate(raw)
        await self.handle(query, data["bot"])

    async def handle(self, query, bot):
        async def answer(text):
            await bot(AnswerWidgetQuery(widget_query_id=query.id, data=text))

        user, chat = query.from_user, query.chat
        if (not chat or chat.type != "private" or chat.id != user.id
                or not query.message_id or not self.access(user.id, user.username)):
            await answer("Доступно только сотрудникам в личке с ботом.")
            return
        if self.get_setting(self.key(chat.id, query.message_id)) != "1":
            await answer("Откройте новую карточку командой /progress2.")
            return
        parts = query.data.split(":")
        if (query.type.get("type") != "data" or len(parts) != 4 or parts[0] != "pw"
                or parts[1] not in self.months() or not parts[2].isdigit()
                or len(parts[2]) > 3 or parts[3] not in {"show", "refresh"}):
            await answer("Некорректный запрос. Откройте /progress2 заново.")
            return
        if user.id in self.busy:
            await answer("Данные уже загружаются…")
            return
        self.busy.add(user.id)
        try:
            # ACK before slow MoySklad reads: Telegram allows only 15 seconds.
            await answer("Обновляю…")
            try:
                state = await self.snapshot(parts[1], int(parts[2]), parts[3] == "refresh")
            except Exception as error:
                logging.warning("Progress widget load failed (%s)", type(error).__name__)
                state = dict(m=parts[1], title=parts[1], prev="", next="", page=0, pages=1,
                             rows=[], error="Не удалось загрузить данные. Нажмите «Обновить».")
            # Permissions may have changed while the data was loading.
            if not self.access(user.id, user.username):
                state = dict(error="Доступ закрыт. Обратитесь к администратору.")
            # Even an unchanged cached result must finish the client's loading state.
            state["r"] = time.time_ns()
            try:
                await bot(EditMessageWidgetData(chat_id=chat.id, message_id=query.message_id,
                                                chat_data=encode_state(state)))
            except TelegramBadRequest as error:
                if "message is not modified" not in str(error).lower():
                    raise
        finally:
            self.busy.discard(user.id)
