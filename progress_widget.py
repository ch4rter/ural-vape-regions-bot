"""Private /progress widget and a narrow Bot API 10.4 adapter for aiogram 3.29.

Protocol: pw:<YYYY-MM>:<page>:<show|refresh>. Financial data is chat_data only.
Docs: https://core.telegram.org/bots/widgets
"""
from __future__ import annotations

import json
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
from aiogram.types import Chat, InlineKeyboardButton, InlineKeyboardMarkup, Message, User
from pydantic import BaseModel, Field

from sales_progress import format_money, month_timing


class SendWidget(TelegramMethod[Message]):
    __api_method__ = "sendWidget"
    __returning__ = Message
    chat_id: int
    html: str
    chat_data: str
    caption: str
    allow_queries: bool = True
    protect_content: bool = True
    reply_markup: InlineKeyboardMarkup
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


def build_state(month, months, title, goal, rows, *, now, updated=None, page=0):
    rows = sorted(rows, key=lambda row: (-row.revenue, row.name.casefold()))
    pages = max(1, (len(rows) + 2) // 3)
    page = min(max(page, 0), pages - 1)
    index = months.index(month)
    state = dict(m=month, title=title, prev=months[index - 1] if index else "",
                 next=months[index + 1] if index + 1 < len(months) else "",
                 page=page, pages=pages, rows=[])
    if goal is None:
        state["error"] = "Цель на этот месяц ещё не установлена."
        return state
    total = sum((row.revenue for row in rows), Decimal(0))
    remaining = max(goal - total, Decimal(0))
    timing, days = month_timing(month, now.date())
    state.update(pct=str((total / goal * 100).quantize(Decimal("0.1"))),
                 total=format_money(total), goal=format_money(goal),
                 left=format_money(remaining), timing=timing.replace("<b>", "").replace("</b>", ""),
                 daily=format_money(remaining / days) if days and remaining else "",
                 updated=updated.strftime("%d.%m %H:%M МСК") if updated else "",
                 rows=[[row.name[:32], format_money(row.revenue)] for row in rows[page * 3:page * 3 + 3]])
    return state


class ProgressWidgets(BaseMiddleware):
    def __init__(self, *, access, months, snapshot, get_setting, set_setting):
        self.access, self.months, self.snapshot = access, months, snapshot
        self.get_setting, self.set_setting = get_setting, set_setting
        self.busy: set[int] = set()
        self.html = Path(__file__).with_name("progress_widget.html").read_text(encoding="utf-8")

    @staticmethod
    def key(chat_id, message_id):
        return f"progress_widget:v1:{chat_id}:{message_id}"

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
                caption="План отдела продаж · /progress\nЕсли виджет не отображается, нажмите «Обычный вид».",
                reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
                    InlineKeyboardButton(text="Обычный вид", callback_data="progress:text")
                ]]), message_thread_id=message.message_thread_id,
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
            await answer("Откройте новую карточку командой /progress.")
            return
        parts = query.data.split(":")
        if (query.type.get("type") != "data" or len(parts) != 4 or parts[0] != "pw"
                or parts[1] not in self.months() or not parts[2].isdigit()
                or len(parts[2]) > 3 or parts[3] not in {"show", "refresh"}):
            await answer("Некорректный запрос. Откройте /progress заново.")
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
