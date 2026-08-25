import asyncio
import html
import logging
import re
from collections import Counter
from datetime import date, datetime, timedelta

from aiogram import F, Router
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, InlineKeyboardButton, InlineKeyboardMarkup, Message

from crm_db import CRMDatabase
from google_crm import CRMClient, CRM_STATUSES, GoogleCRM, PRODUCT_COLUMNS


router = Router(name="crm_beta")
crm: GoogleCRM | None = None
database: CRMDatabase | None = None
owner_ids: set[int] = set()
TASKS_PER_PAGE = 10
CALL_LIST_PER_PAGE = 12


class CRMState(StatesGroup):
    search = State()
    note = State()
    product_reason = State()
    task_text = State()
    task_date = State()
    filter_region = State()


def configure_crm(service: GoogleCRM | None, db: CRMDatabase, owners: set[int]) -> None:
    global crm, database, owner_ids
    crm, database, owner_ids = service, db, owners


def allowed(user_id: int | None) -> bool:
    return bool(user_id and user_id in owner_ids)


async def guard(event: CallbackQuery | Message) -> bool:
    user_id = event.from_user.id if event.from_user else None
    if not allowed(user_id):
        if isinstance(event, CallbackQuery):
            await event.answer("CRM-бета пока доступна только владельцу.", show_alert=True)
        return False
    if crm is None:
        text = "⚠️ CRM временно недоступна. Проверьте подключение к Google Таблице."
        if isinstance(event, CallbackQuery):
            await event.answer(text, show_alert=True)
        else:
            await event.answer(text)
        return False
    return True


def nav(back: str = "crm:menu") -> list[InlineKeyboardButton]:
    return [
        InlineKeyboardButton(text="⬅️", callback_data=back),
        InlineKeyboardButton(text="🏠", callback_data="main:menu"),
    ]


def crm_menu_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🔎 Найти клиента", callback_data="crm:search"),
            InlineKeyboardButton(text="📋 Мой список", callback_data="crm:call_list"),
        ],
        [
            InlineKeyboardButton(text="🎯 Подобрать", callback_data="crm:pick"),
            InlineKeyboardButton(text="✅ Задачи", callback_data="crm:tasks"),
        ],
        [InlineKeyboardButton(text="📝 Итоги дня", callback_data="crm:report")],
        [InlineKeyboardButton(text="🏠", callback_data="main:menu")],
    ])


async def render(message: Message, text: str, markup: InlineKeyboardMarkup) -> None:
    try:
        await message.edit_text(text, reply_markup=markup)
    except Exception:
        await message.answer(text, reply_markup=markup)


def product_icon(state: str) -> str:
    return {"buy": "🟢", "no": "🔴", "unknown": "⚪"}.get(state, "⚪")


def short_note(value: str, limit: int = 700) -> str:
    if not value:
        return "—"
    return value if len(value) <= limit else f"…{value[-limit:]}"


def client_text(client: CRMClient) -> str:
    product_lines = []
    for key, (_, label) in PRODUCT_COLUMNS.items():
        line = f"{product_icon(client.products[key])} {label}"
        if client.products[key] == "no":
            reason = client.product_notes.get(key, "").strip()
            line += f" — {html.escape(reason)}" if reason else " — <i>причина не заполнена</i>"
        elif client.product_notes.get(key, "").strip():
            line += f" · {html.escape(client.product_notes[key])}"
        product_lines.append(line)
    products = "\n".join(product_lines)
    contact = " · ".join(value for value in (client.contact_name, client.phone, client.telegram) if value) or "—"
    return (
        f"👤 <b>{html.escape(client.name)}</b>\n"
        f"📍 {html.escape(client.region or '—')} · {html.escape(client.client_type or '—')} · "
        f"ТТ: <b>{html.escape(client.outlets or '—')}</b>\n"
        f"Статус: <b>{html.escape(client.status or 'не заполнен')}</b>\n"
        f"Контакт: {html.escape(contact)}\n"
        f"Последний заказ: <b>{html.escape(client.order_date or '—')}</b>\n"
        f"Последний звонок: <b>{html.escape(client.call_date or '—')}</b>\n\n"
        f"<b>Товарные группы</b>\n{products}\n\n"
        f"<b>Последнее из примечаний</b>\n{html.escape(short_note(client.note))}"
    )


def client_keyboard(client: CRMClient, in_list: bool) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton(text="📝 Результат", callback_data="crm:add_note"),
            InlineKeyboardButton(text="📅 Задача", callback_data="crm:add_task"),
        ],
        [
            InlineKeyboardButton(text="🔄 Статус", callback_data="crm:statuses"),
            InlineKeyboardButton(text="🛍 Товары", callback_data="crm:products"),
        ],
        [InlineKeyboardButton(text="🧾 Заказ сегодня", callback_data="crm:order_today")],
    ]
    if in_list:
        rows.append([
            InlineKeyboardButton(text="➖ Из списка", callback_data="crm:list_remove"),
            InlineKeyboardButton(text="➡️ Следующий", callback_data="crm:list_next"),
        ])
    else:
        rows.append([InlineKeyboardButton(text="➕ В список обзвона", callback_data="crm:list_add")])
    if client.telegram.startswith("@") and re.fullmatch(r"@[A-Za-z0-9_]{5,32}", client.telegram):
        rows.append([InlineKeyboardButton(text="✈️ Открыть Telegram", url=f"https://t.me/{client.telegram[1:]}")])
    rows.append(nav("crm:search"))
    return InlineKeyboardMarkup(inline_keyboard=rows)


async def set_current(state: FSMContext, client: CRMClient) -> None:
    await state.update_data(crm_row=client.row, crm_identity=client.identity)


async def current_client(state: FSMContext) -> CRMClient | None:
    data = await state.get_data()
    row = int(data.get("crm_row", 0))
    identity = data.get("crm_identity", "")
    if not row:
        return None
    return await asyncio.to_thread(crm.client, row, identity)


async def show_client(message: Message, state: FSMContext, client: CRMClient) -> None:
    await set_current(state, client)
    listed = any(
        row["client_identity"] == client.identity
        for row in database.call_list(message.chat.id)
    )
    await render(message, client_text(client), client_keyboard(client, listed))


@router.callback_query(F.data == "crm:menu")
async def open_menu(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.clear()
    due_tasks = database.tasks(callback.from_user.id)
    all_tasks = database.tasks(callback.from_user.id, include_future=True)
    call_count = len(database.call_list(callback.from_user.id))
    overdue = sum(task.due_date < date.today().isoformat() for task in due_tasks)
    await render(
        callback.message,
        "📇 <b>CRM · Активные клиенты</b>\n\n"
        f"В списке обзвона: <b>{call_count}</b>\n"
        f"Активных задач: <b>{len(all_tasks)}</b> · на сегодня и просроченных: <b>{len(due_tasks)}</b>"
        + (f" · просрочено <b>{overdue}</b>" if overdue else "")
        + "\n\nВыберите действие:",
        crm_menu_keyboard(),
    )
    await callback.answer()


@router.callback_query(F.data == "crm:search")
async def start_search(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.set_state(CRMState.search)
    await render(
        callback.message,
        "🔎 <b>Поиск клиента</b>\n\n"
        "Введите название компании, имя, регион, телефон или Telegram. Точное совпадение не требуется.",
        InlineKeyboardMarkup(inline_keyboard=[nav()]),
    )
    await callback.answer()


@router.message(CRMState.search, F.text)
async def search_clients(message: Message, state: FSMContext) -> None:
    if not await guard(message):
        return
    query = message.text.strip()
    if len(query) < 2:
        await message.answer("Введите хотя бы два символа.")
        return
    status = await message.answer("⏳ Ищу клиента…")
    try:
        results = await asyncio.to_thread(crm.search, query)
    except Exception:
        logging.exception("Ошибка поиска в Google CRM")
        await status.edit_text("⚠️ Не удалось прочитать Google Таблицу. Попробуйте ещё раз.")
        return
    await state.update_data(crm_search=[{"row": item.row, "identity": item.identity} for item in results])
    if not results:
        await status.edit_text(
            "Ничего не найдено. Попробуйте название короче или используйте имя/Telegram.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[nav("crm:search")]),
        )
        return
    rows = [[InlineKeyboardButton(
        text=f"{item.name} · {item.region}"[:64], callback_data=f"crm:result:{index}"
    )] for index, item in enumerate(results)]
    rows.append(nav("crm:search"))
    await status.edit_text(
        f"🔎 <b>Найдено: {len(results)}</b>\n\nВыберите клиента:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )


@router.callback_query(F.data.startswith("crm:result:"))
async def open_result(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    try:
        index = int(callback.data.rsplit(":", 1)[1])
        item = (await state.get_data())["crm_search"][index]
        client = await asyncio.to_thread(crm.client, item["row"], item["identity"])
    except Exception:
        client = None
    if not client:
        await callback.answer("Карточка изменилась. Выполните поиск ещё раз.", show_alert=True)
        return
    await show_client(callback.message, state, client)
    await callback.answer()


@router.callback_query(F.data == "crm:add_note")
async def start_note(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    await state.set_state(CRMState.note)
    await callback.message.edit_text(
        f"📝 <b>{html.escape(client.name)}</b>\n\n"
        "Напишите результат общения одним сообщением. Он будет добавлен в конец примечания с датой и временем, а дата звонка обновится автоматически.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[nav("crm:search")]),
    )
    await callback.answer()


@router.message(CRMState.note, F.text)
async def save_note(message: Message, state: FSMContext) -> None:
    if not await guard(message):
        return
    text = message.text.strip()
    if len(text) < 3 or len(text) > 3000:
        await message.answer("Запись должна содержать от 3 до 3000 символов.")
        return
    client = await current_client(state)
    if not client:
        await message.answer("Карточка изменилась. Найдите клиента повторно.")
        return
    waiting = await message.answer("⏳ Сохраняю в Google Таблицу…")
    try:
        await asyncio.to_thread(crm.append_note, client.row, client.identity, text, datetime.now())
        database.add_event(message.from_user.id, client.identity, client.name, "note", text)
        fresh = await asyncio.to_thread(crm.client, client.row, client.identity)
        await state.set_state(None)
        await waiting.delete()
        await show_client(message, state, fresh or client)
    except Exception as error:
        logging.exception("Не удалось сохранить CRM-примечание")
        await waiting.edit_text(f"⚠️ Не удалось сохранить запись: {html.escape(str(error))}")


@router.callback_query(F.data == "crm:statuses")
async def statuses(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    rows = [[InlineKeyboardButton(
        text=("✅ " if status.casefold() == client.status.casefold() else "") + status,
        callback_data=f"crm:status:{index}",
    )] for index, status in enumerate(CRM_STATUSES)]
    rows.append(nav("crm:card"))
    await callback.message.edit_text(
        f"🔄 <b>{html.escape(client.name)}</b>\n\nВыберите новый статус:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("crm:status:"))
async def set_status(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    try:
        status = CRM_STATUSES[int(callback.data.rsplit(":", 1)[1])]
    except Exception:
        status = ""
    if not client or not status:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    await callback.answer("Сохраняю…")
    try:
        await asyncio.to_thread(crm.update_status, client.row, client.identity, status)
        database.add_event(
            callback.from_user.id, client.identity, client.name, "status",
            f"Статус: {client.status or 'не заполнен'} → {status}",
        )
        fresh = await asyncio.to_thread(crm.client, client.row, client.identity)
        await show_client(callback.message, state, fresh or client)
    except Exception as error:
        logging.exception("Не удалось обновить CRM-статус")
        await callback.message.answer(f"⚠️ Не удалось обновить статус: {html.escape(str(error))}")


@router.callback_query(F.data == "crm:products")
async def products(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    rows = [[InlineKeyboardButton(
        text=f"{product_icon(client.products[key])} {label}", callback_data=f"crm:product:{key}"
    )] for key, (_, label) in PRODUCT_COLUMNS.items()]
    rows.append(nav("crm:card"))
    await callback.message.edit_text(
        f"🛍 <b>{html.escape(client.name)}</b>\n\n"
        "🟢 покупает · 🔴 не покупает · ⚪ не заполнено\n\nВыберите товарную группу:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("crm:product:"))
async def choose_product(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    key = callback.data.rsplit(":", 1)[1]
    if key not in PRODUCT_COLUMNS:
        await callback.answer("Неизвестная группа.", show_alert=True)
        return
    await state.update_data(crm_product=key)
    label = PRODUCT_COLUMNS[key][1]
    await callback.message.edit_text(
        f"🛍 <b>{html.escape(label)}</b>\n\nКак клиент работает с этой группой?",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🟢 Покупает", callback_data="crm:product_set:buy")],
            [InlineKeyboardButton(text="🔴 Не покупает", callback_data="crm:product_set:no")],
            [InlineKeyboardButton(text="⚪ Не заполнено", callback_data="crm:product_set:unknown")],
            nav("crm:products"),
        ]),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("crm:product_set:"))
async def set_product_state(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    product_state = callback.data.rsplit(":", 1)[1]
    data = await state.get_data()
    product = data.get("crm_product", "")
    client = await current_client(state)
    if not client or product not in PRODUCT_COLUMNS:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    if product_state == "no":
        await state.set_state(CRMState.product_reason)
        await callback.message.edit_text(
            f"🔴 <b>{html.escape(PRODUCT_COLUMNS[product][1])}: не покупает</b>\n\n"
            "Обязательно напишите причину. Она будет сохранена внутри красной ячейки.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[nav("crm:products")]),
        )
        await callback.answer()
        return
    await callback.answer("Сохраняю…")
    await save_product(callback.message, state, callback.from_user.id, client, product, product_state, "")


@router.message(CRMState.product_reason, F.text)
async def save_product_reason(message: Message, state: FSMContext) -> None:
    if not await guard(message):
        return
    reason = message.text.strip()
    if len(reason) < 3 or len(reason) > 500:
        await message.answer("Причина должна содержать от 3 до 500 символов.")
        return
    data = await state.get_data()
    product = data.get("crm_product", "")
    client = await current_client(state)
    if not client or product not in PRODUCT_COLUMNS:
        await message.answer("Откройте клиента повторно.")
        return
    await save_product(message, state, message.from_user.id, client, product, "no", reason)


async def save_product(
    message: Message, state: FSMContext, user_id: int, client: CRMClient,
    product: str, product_state: str, reason: str,
) -> None:
    waiting = await message.answer("⏳ Обновляю товарную группу…")
    try:
        await asyncio.to_thread(
            crm.set_product, client.row, client.identity, product, product_state, reason
        )
        label = PRODUCT_COLUMNS[product][1]
        state_label = {"buy": "покупает", "no": f"не покупает — {reason}", "unknown": "не заполнено"}[product_state]
        database.add_event(user_id, client.identity, client.name, "product", f"{label}: {state_label}")
        fresh = await asyncio.to_thread(crm.client, client.row, client.identity)
        await state.set_state(None)
        await waiting.delete()
        await show_client(message, state, fresh or client)
    except Exception as error:
        logging.exception("Не удалось обновить товарную группу CRM")
        await waiting.edit_text(f"⚠️ Не удалось сохранить: {html.escape(str(error))}")


@router.callback_query(F.data == "crm:order_today")
async def order_today(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    await callback.answer("Сохраняю…")
    try:
        await asyncio.to_thread(crm.mark_order_today, client.row, client.identity, datetime.now())
        database.add_event(callback.from_user.id, client.identity, client.name, "order", "Заказ сегодня")
        fresh = await asyncio.to_thread(crm.client, client.row, client.identity)
        await show_client(callback.message, state, fresh or client)
    except Exception as error:
        logging.exception("Не удалось отметить заказ в CRM")
        await callback.message.answer(f"⚠️ Не удалось отметить заказ: {html.escape(str(error))}")


@router.callback_query(F.data == "crm:add_task")
async def start_task(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    await state.set_state(CRMState.task_text)
    await callback.message.edit_text(
        f"📅 <b>Новая задача · {html.escape(client.name)}</b>\n\nЧто нужно сделать?",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[nav("crm:search")]),
    )
    await callback.answer()


@router.message(CRMState.task_text, F.text)
async def task_text(message: Message, state: FSMContext) -> None:
    if not await guard(message):
        return
    text = message.text.strip()
    if len(text) < 3 or len(text) > 500:
        await message.answer("Описание должно содержать от 3 до 500 символов.")
        return
    await state.update_data(crm_task_text=text)
    await state.set_state(CRMState.task_date)
    await message.answer(
        "Когда напомнить? Напишите <code>сегодня</code>, <code>завтра</code>, "
        "дату <code>07.08</code> или <code>07.08.2026</code>.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[nav("crm:search")]),
    )


def parse_task_date(value: str) -> date | None:
    normalized = value.strip().casefold()
    if normalized == "сегодня":
        return date.today()
    if normalized == "завтра":
        return date.today() + timedelta(days=1)
    try:
        return datetime.strptime(normalized, "%d.%m.%Y").date()
    except ValueError:
        pass
    try:
        day, month = (int(part) for part in normalized.split("."))
        result = date(date.today().year, month, day)
        if result < date.today():
            result = result.replace(year=result.year + 1)
        return result
    except (ValueError, TypeError):
        pass
    return None


@router.message(CRMState.task_date, F.text)
async def task_date(message: Message, state: FSMContext) -> None:
    if not await guard(message):
        return
    due = parse_task_date(message.text)
    if not due:
        await message.answer("Не понял дату. Например: <code>завтра</code> или <code>07.08</code>.")
        return
    client = await current_client(state)
    data = await state.get_data()
    if not client:
        await message.answer("Откройте клиента повторно.")
        return
    task = database.add_task(
        message.from_user.id, client.identity, client.name, due, data["crm_task_text"]
    )
    database.add_event(
        message.from_user.id, client.identity, client.name, "task",
        f"Задача на {due.strftime('%d.%m.%Y')}: {task.text}",
    )
    await state.set_state(None)
    await message.answer(
        f"✅ Задача сохранена на <b>{due.strftime('%d.%m.%Y')}</b>.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[nav("crm:menu")]),
    )


@router.callback_query(F.data == "crm:tasks")
async def tasks(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.clear()
    await render_tasks(callback, 0)


@router.callback_query(F.data.startswith("crm:tasks_page:"))
async def tasks_page(callback: CallbackQuery) -> None:
    if not await guard(callback):
        return
    await render_tasks(callback, int(callback.data.rsplit(":", 1)[1]))


async def render_tasks(callback: CallbackQuery, page: int) -> None:
    items = database.tasks(callback.from_user.id, include_future=True)
    if not items:
        await render(
            callback.message, "✅ <b>Все задачи</b>\n\nАктивных задач пока нет.",
            InlineKeyboardMarkup(inline_keyboard=[nav()]),
        )
        await callback.answer()
        return
    page_count = max(1, (len(items) + TASKS_PER_PAGE - 1) // TASKS_PER_PAGE)
    page = max(0, min(page, page_count - 1))
    start = page * TASKS_PER_PAGE
    shown = items[start:start + TASKS_PER_PAGE]
    today = date.today().isoformat()
    overdue = sum(task.due_date < today for task in items)
    due_today = sum(task.due_date == today for task in items)
    future = sum(task.due_date > today for task in items)
    lines = [
        "✅ <b>Все активные задачи</b>",
        f"🔴 Просрочено: <b>{overdue}</b> · 📅 Сегодня: <b>{due_today}</b> · 🔵 Позже: <b>{future}</b>",
        f"Страница <b>{page + 1}</b> из <b>{page_count}</b>",
    ]
    rows = []
    for index, task in enumerate(shown, start + 1):
        icon = "🔴" if task.due_date < today else "📅" if task.due_date == today else "🔵"
        lines.extend(["", f"<b>{index}. {html.escape(task.client_name)}</b>",
                      f"{icon} {datetime.fromisoformat(task.due_date).strftime('%d.%m.%Y')} · {html.escape(task.text)}"])
        rows.append([InlineKeyboardButton(
            text=f"✅ Выполнено · {task.client_name}"[:64],
            callback_data=f"crm:task_done:{task.id}", style="success",
        )])
    pagination = []
    if page > 0:
        pagination.append(InlineKeyboardButton(text="⬅️", callback_data=f"crm:tasks_page:{page - 1}"))
    if page + 1 < page_count:
        pagination.append(InlineKeyboardButton(text="➡️", callback_data=f"crm:tasks_page:{page + 1}"))
    if pagination:
        rows.append(pagination)
    rows.append(nav())
    await render(callback.message, "\n".join(lines), InlineKeyboardMarkup(inline_keyboard=rows))
    await callback.answer()


@router.callback_query(F.data.startswith("crm:task_done:"))
async def task_done(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    task_id = int(callback.data.rsplit(":", 1)[1])
    database.complete_task(callback.from_user.id, task_id)
    await render_tasks(callback, 0)


@router.callback_query(F.data == "crm:list_add")
async def list_add(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента повторно.", show_alert=True)
        return
    added = database.add_to_call_list(callback.from_user.id, client.identity, client.name)
    await callback.answer("Добавлен в список" if added else "Уже находится в списке")
    await show_client(callback.message, state, client)


@router.callback_query(F.data == "crm:list_remove")
async def list_remove(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if client:
        database.remove_from_call_list(callback.from_user.id, client.identity)
    await callback.answer("Удалён из списка")
    if client:
        await show_client(callback.message, state, client)


@router.callback_query(F.data == "crm:call_list")
async def call_list(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.clear()
    items = database.call_list(callback.from_user.id)
    await state.update_data(crm_call_identities=[row["client_identity"] for row in items])
    await render_call_list(callback, items, 0)


@router.callback_query(F.data.startswith("crm:call_page:"))
async def call_list_page(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    items = database.call_list(callback.from_user.id)
    await state.update_data(crm_call_identities=[row["client_identity"] for row in items])
    await render_call_list(callback, items, int(callback.data.rsplit(":", 1)[1]))


async def render_call_list(callback: CallbackQuery, items, page: int) -> None:
    page_count = max(1, (len(items) + CALL_LIST_PER_PAGE - 1) // CALL_LIST_PER_PAGE)
    page = max(0, min(page, page_count - 1))
    start = page * CALL_LIST_PER_PAGE
    rows = [[InlineKeyboardButton(
        text=f"{index + 1}. {row['client_name']}"[:64], callback_data=f"crm:call:{index}"
    )] for index, row in enumerate(items[start:start + CALL_LIST_PER_PAGE], start)]
    pagination = []
    if page > 0:
        pagination.append(InlineKeyboardButton(text="⬅️", callback_data=f"crm:call_page:{page - 1}"))
    if page + 1 < page_count:
        pagination.append(InlineKeyboardButton(text="➡️", callback_data=f"crm:call_page:{page + 1}"))
    if pagination:
        rows.append(pagination)
    if items:
        rows.append([InlineKeyboardButton(text="🗑 Очистить список", callback_data="crm:list_clear", style="danger")])
    rows.append(nav())
    text = (
        f"📋 <b>Мой список обзвона</b>\n\nКлиентов: <b>{len(items)}</b>\n"
        f"Страница <b>{page + 1}</b> из <b>{page_count}</b>\n"
        "Открывайте карточки по порядку и фиксируйте результат."
        if items else "📋 <b>Мой список обзвона</b>\n\nСписок пока пуст. Добавьте клиентов через поиск или подборку."
    )
    await render(callback.message, text, InlineKeyboardMarkup(inline_keyboard=rows))
    await callback.answer()


@router.callback_query(F.data.startswith("crm:call:"))
async def open_call_item(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    data = await state.get_data()
    try:
        identity = data["crm_call_identities"][int(callback.data.rsplit(":", 1)[1])]
        client = next(item for item in await asyncio.to_thread(crm.clients) if item.identity == identity)
    except Exception:
        client = None
    if not client:
        await callback.answer("Клиент не найден в таблице.", show_alert=True)
        return
    await show_client(callback.message, state, client)
    await callback.answer()


@router.callback_query(F.data == "crm:list_next")
async def list_next(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    items = database.call_list(callback.from_user.id)
    identities = [row["client_identity"] for row in items]
    try:
        next_identity = identities[identities.index(client.identity) + 1]
    except (ValueError, IndexError, AttributeError):
        await callback.answer("Это последний клиент в списке.", show_alert=True)
        return
    next_client = next(
        (item for item in await asyncio.to_thread(crm.clients) if item.identity == next_identity), None
    )
    if not next_client:
        await callback.answer("Следующий клиент не найден в таблице.", show_alert=True)
        return
    await show_client(callback.message, state, next_client)
    await callback.answer()


@router.callback_query(F.data == "crm:list_clear")
async def list_clear(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    database.clear_call_list(callback.from_user.id)
    await call_list(callback, state)


def parse_sheet_date(value: str) -> date | None:
    for pattern in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(value.strip(), pattern).date()
        except (ValueError, AttributeError):
            pass
    return None


@router.callback_query(F.data == "crm:pick")
async def pick_menu(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.clear()
    await render(
        callback.message,
        "🎯 <b>Подобрать клиентов</b>\n\nВыберите готовое условие для первой бета-версии:",
        InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⚙️ Настроить по характеристикам", callback_data="crm:filter")],
            [InlineKeyboardButton(text="📞 Не звонили 30 дней", callback_data="crm:pick_do:old_call")],
            [InlineKeyboardButton(text="🛑 Перестали заказывать", callback_data="crm:pick_do:stopped")],
            [InlineKeyboardButton(text="⏳ Ждём заказ", callback_data="crm:pick_do:waiting")],
            [InlineKeyboardButton(text="⚪ Не заполнены товары", callback_data="crm:pick_do:unknown")],
            [InlineKeyboardButton(text="🔴 Нет причины отказа", callback_data="crm:pick_do:missing_reason")],
            [InlineKeyboardButton(text="🎲 Случайные 10", callback_data="crm:pick_do:random")],
            nav(),
        ]),
    )
    await callback.answer()


def default_filter() -> dict:
    return {
        "kind": "all",
        "scale": "any",
        "region": "",
        "statuses": [],
        "products": [],
        "product_mode": "any",
        "product_state": "buy",
    }


def filter_keyboard(settings: dict) -> InlineKeyboardMarkup:
    kind = settings.get("kind", "all")
    scale = settings.get("scale", "any")
    products = set(settings.get("products", []))
    mode = settings.get("product_mode", "any")
    product_state = settings.get("product_state", "buy")
    statuses = settings.get("statuses", [])
    region = settings.get("region", "")
    mark = lambda selected, label: f"✅ {label}" if selected else label
    rows = [
        [
            InlineKeyboardButton(text=mark(kind == "all", "Все"), callback_data="crm:filter_kind:all"),
            InlineKeyboardButton(text=mark(kind == "retail", "Розница"), callback_data="crm:filter_kind:retail"),
            InlineKeyboardButton(text=mark(kind == "wholesale", "Опт"), callback_data="crm:filter_kind:wholesale"),
        ],
        [
            InlineKeyboardButton(text=mark(scale == "any", "Любое кол-во"), callback_data="crm:filter_scale:any"),
            InlineKeyboardButton(text=mark(scale == "1_3", "1–3 ТТ"), callback_data="crm:filter_scale:1_3"),
        ],
        [
            InlineKeyboardButton(text=mark(scale == "4_10", "4–10 ТТ"), callback_data="crm:filter_scale:4_10"),
            InlineKeyboardButton(text=mark(scale == "11_plus", "11+ ТТ"), callback_data="crm:filter_scale:11_plus"),
        ],
        [
            InlineKeyboardButton(
                text=f"📍 Регион{' · ' + region[:18] if region else ''}", callback_data="crm:filter_region"
            ),
            InlineKeyboardButton(
                text=f"🔄 Статусы · {len(statuses)}", callback_data="crm:filter_statuses"
            ),
        ],
    ]
    product_buttons = [InlineKeyboardButton(
        text=mark(key in products, label), callback_data=f"crm:filter_product:{key}"
    ) for key, (_, label) in PRODUCT_COLUMNS.items()]
    rows.extend([product_buttons[index:index + 2] for index in range(0, len(product_buttons), 2)])
    rows.append([
        InlineKeyboardButton(
            text=mark(product_state == "buy", "🟢 Покупает"), callback_data="crm:filter_product_state:buy"
        ),
        InlineKeyboardButton(
            text=mark(product_state == "no", "🔴 Не покупает"), callback_data="crm:filter_product_state:no"
        ),
    ])
    rows.append([InlineKeyboardButton(
        text=mark(product_state == "unknown", "⚪ Не заполнено"),
        callback_data="crm:filter_product_state:unknown",
    )])
    rows.append([
        InlineKeyboardButton(
            text=mark(mode == "any", "Хотя бы одна"), callback_data="crm:filter_mode:any"
        ),
        InlineKeyboardButton(
            text=mark(mode == "all", "Все выбранные"), callback_data="crm:filter_mode:all"
        ),
    ])
    rows.append([InlineKeyboardButton(text="🔎 Показать клиентов", callback_data="crm:filter_apply")])
    rows.append(nav("crm:pick"))
    return InlineKeyboardMarkup(inline_keyboard=rows)


def filter_description(settings: dict) -> str:
    kind_names = {"all": "все типы", "retail": "розница", "wholesale": "опт и опт+розница"}
    scale_names = {"any": "любое количество ТТ", "1_3": "1–3 ТТ", "4_10": "4–10 ТТ", "11_plus": "11+ ТТ"}
    selected = [PRODUCT_COLUMNS[key][1] for key in settings.get("products", []) if key in PRODUCT_COLUMNS]
    products = ", ".join(selected) if selected else "любые товарные группы"
    mode = "хотя бы одну" if settings.get("product_mode") == "any" else "все выбранные"
    selected_statuses = []
    for value in settings.get("statuses", []):
        selected_statuses.append("не заполнен" if value == "__empty__" else value)
    statuses = ", ".join(selected_statuses) if selected_statuses else "любые"
    product_state = {
        "buy": "покупает", "no": "не покупает", "unknown": "не заполнено",
    }.get(settings.get("product_state"), "покупает")
    return (
        "⚙️ <b>Подбор по характеристикам</b>\n\n"
        f"Регион: <b>{html.escape(settings.get('region') or 'любой')}</b>\n"
        f"Тип клиента: <b>{kind_names.get(settings.get('kind'), 'все типы')}</b>\n"
        f"Масштаб: <b>{scale_names.get(settings.get('scale'), 'любое количество ТТ')}</b>\n"
        f"Статус: <b>{html.escape(statuses)}</b>\n"
        f"Товары ({product_state}): <b>{html.escape(products)}</b>"
        + (f" · условие «{mode}»" if selected else "")
        + "\n\nМожно выбрать несколько товарных групп."
    )


async def show_filter(callback: CallbackQuery, state: FSMContext, reset: bool = False) -> None:
    data = await state.get_data()
    settings = default_filter() if reset else data.get("crm_filter", default_filter())
    await state.update_data(crm_filter=settings)
    await state.set_state(None)
    await render(callback.message, filter_description(settings), filter_keyboard(settings))
    await callback.answer()


@router.callback_query(F.data == "crm:filter")
async def open_filter(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.clear()
    await show_filter(callback, state, reset=True)


@router.callback_query(F.data.startswith("crm:filter_kind:"))
async def set_filter_kind(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["kind"] = callback.data.rsplit(":", 1)[1]
    await state.update_data(crm_filter=settings)
    await show_filter(callback, state)


@router.callback_query(F.data.startswith("crm:filter_scale:"))
async def set_filter_scale(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["scale"] = callback.data.rsplit(":", 1)[1]
    await state.update_data(crm_filter=settings)
    await show_filter(callback, state)


@router.callback_query(F.data == "crm:filter_region")
async def ask_filter_region(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    rows = []
    if settings.get("region"):
        rows.append([InlineKeyboardButton(text="🗑 Сбросить регион", callback_data="crm:filter_region_clear", style="danger")])
    rows.append(nav("crm:filter"))
    await state.set_state(CRMState.filter_region)
    await callback.message.edit_text(
        "📍 <b>Регион клиента</b>\n\n"
        "Введите город, область или часть названия региона. Точное совпадение не требуется.",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows),
    )
    await callback.answer()


@router.message(CRMState.filter_region, F.text)
async def save_filter_region(message: Message, state: FSMContext) -> None:
    if not await guard(message):
        return
    value = message.text.strip()
    if len(value) < 2 or len(value) > 100:
        await message.answer("Введите от 2 до 100 символов.")
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["region"] = value
    await state.update_data(crm_filter=settings)
    await state.set_state(None)
    await render(message, filter_description(settings), filter_keyboard(settings))


@router.callback_query(F.data == "crm:filter_region_clear")
async def clear_filter_region(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["region"] = ""
    await state.update_data(crm_filter=settings)
    await state.set_state(None)
    await show_filter(callback, state)


def status_filter_keyboard(settings: dict) -> InlineKeyboardMarkup:
    selected = set(settings.get("statuses", []))
    choices = list(CRM_STATUSES) + ["__empty__"]
    rows = []
    for index, value in enumerate(choices):
        label = "Не заполнен" if value == "__empty__" else value
        if value in selected:
            label = f"✅ {label}"
        rows.append([InlineKeyboardButton(text=label, callback_data=f"crm:filter_status:{index}")])
    if selected:
        rows.append([InlineKeyboardButton(text="🗑 Сбросить статусы", callback_data="crm:filter_status_clear", style="danger")])
    rows.append(nav("crm:filter"))
    return InlineKeyboardMarkup(inline_keyboard=rows)


@router.callback_query(F.data == "crm:filter_statuses")
async def open_filter_statuses(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    await callback.message.edit_text(
        "🔄 <b>Статус работы</b>\n\n"
        "Отметьте один или несколько статусов. Клиент подойдёт, если у него установлен любой из выбранных.",
        reply_markup=status_filter_keyboard(settings),
    )
    await callback.answer()


@router.callback_query(F.data.startswith("crm:filter_status:"))
async def toggle_filter_status(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    choices = list(CRM_STATUSES) + ["__empty__"]
    try:
        value = choices[int(callback.data.rsplit(":", 1)[1])]
    except (ValueError, IndexError):
        await callback.answer("Неизвестный статус.", show_alert=True)
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    selected = list(settings.get("statuses", []))
    selected.remove(value) if value in selected else selected.append(value)
    settings["statuses"] = selected
    await state.update_data(crm_filter=settings)
    await callback.message.edit_reply_markup(reply_markup=status_filter_keyboard(settings))
    await callback.answer()


@router.callback_query(F.data == "crm:filter_status_clear")
async def clear_filter_statuses(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["statuses"] = []
    await state.update_data(crm_filter=settings)
    await callback.message.edit_reply_markup(reply_markup=status_filter_keyboard(settings))
    await callback.answer("Статусы сброшены.")


@router.callback_query(F.data.startswith("crm:filter_product:"))
async def toggle_filter_product(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    key = callback.data.rsplit(":", 1)[1]
    settings = (await state.get_data()).get("crm_filter", default_filter())
    selected = list(settings.get("products", []))
    if key in selected:
        selected.remove(key)
    elif key in PRODUCT_COLUMNS:
        selected.append(key)
    settings["products"] = selected
    await state.update_data(crm_filter=settings)
    await show_filter(callback, state)


@router.callback_query(F.data.startswith("crm:filter_mode:"))
async def set_filter_mode(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["product_mode"] = callback.data.rsplit(":", 1)[1]
    await state.update_data(crm_filter=settings)
    await show_filter(callback, state)


@router.callback_query(F.data.startswith("crm:filter_product_state:"))
async def set_filter_product_state(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    value = callback.data.rsplit(":", 1)[1]
    if value not in {"buy", "no", "unknown"}:
        await callback.answer("Неизвестное состояние.", show_alert=True)
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    settings["product_state"] = value
    await state.update_data(crm_filter=settings)
    await show_filter(callback, state)


def numeric_outlets(value: str) -> int | None:
    match = re.search(r"\d+(?:[.,]\d+)?", value or "")
    return int(float(match.group(0).replace(",", "."))) if match else None


def matches_filter(client: CRMClient, settings: dict) -> bool:
    region_query = re.sub(r"\s+", " ", settings.get("region", "").casefold().replace("ё", "е")).strip()
    client_region = re.sub(r"\s+", " ", client.region.casefold().replace("ё", "е")).strip()
    if region_query and region_query not in client_region:
        return False
    client_type = client.client_type.casefold().replace("ё", "е")
    kind = settings.get("kind", "all")
    if kind == "retail" and "розниц" not in client_type:
        return False
    if kind == "wholesale" and "опт" not in client_type and "опт" not in client.outlets.casefold():
        return False
    outlets = numeric_outlets(client.outlets)
    scale = settings.get("scale", "any")
    if scale == "1_3" and (outlets is None or not 1 <= outlets <= 3):
        return False
    if scale == "4_10" and (outlets is None or not 4 <= outlets <= 10):
        return False
    if scale == "11_plus" and (outlets is None or outlets < 11):
        return False
    statuses = settings.get("statuses", [])
    if statuses:
        current_status = client.status.strip().casefold().replace("ё", "е")
        allowed_statuses = {
            value.casefold().replace("ё", "е") for value in statuses if value != "__empty__"
        }
        if current_status not in allowed_statuses and not ("__empty__" in statuses and not current_status):
            return False
    products = [key for key in settings.get("products", []) if key in PRODUCT_COLUMNS]
    if products:
        product_state = settings.get("product_state", "buy")
        matches = [client.products.get(key) == product_state for key in products]
        if settings.get("product_mode") == "all" and not all(matches):
            return False
        if settings.get("product_mode") != "all" and not any(matches):
            return False
    return True


@router.callback_query(F.data == "crm:filter_apply")
async def apply_filter(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    settings = (await state.get_data()).get("crm_filter", default_filter())
    await callback.answer("Подбираю клиентов…")
    clients = await asyncio.to_thread(crm.clients)
    selected = [client for client in clients if matches_filter(client, settings)]
    await state.update_data(crm_pick=[{"identity": item.identity, "name": item.name} for item in selected])
    preview = "\n".join(f"• {html.escape(item.name)}" for item in selected[:15])
    if len(selected) > 15:
        preview += f"\n…и ещё {len(selected) - 15}"
    await render(
        callback.message,
        filter_description(settings)
        + f"\n\n<b>Найдено: {len(selected)}</b>\n\n{preview or 'Никого не найдено.'}",
        InlineKeyboardMarkup(inline_keyboard=(
            [[InlineKeyboardButton(text="➕ Добавить всех в мой список", callback_data="crm:pick_add")]]
            if selected else []
        ) + [nav("crm:filter")]),
    )


@router.callback_query(F.data.startswith("crm:pick_do:"))
async def pick_clients(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    mode = callback.data.rsplit(":", 1)[1]
    await callback.answer("Формирую список…")
    clients = await asyncio.to_thread(crm.clients)
    today = date.today()
    if mode == "old_call":
        selected = [item for item in clients if not parse_sheet_date(item.call_date) or (today - parse_sheet_date(item.call_date)).days >= 30]
    elif mode == "stopped":
        selected = [item for item in clients if "перестал" in item.status.casefold()]
    elif mode == "waiting":
        selected = [item for item in clients if "ждем заказ" in item.status.casefold().replace("ё", "е")]
    elif mode == "unknown":
        selected = [item for item in clients if any(value == "unknown" for value in item.products.values())]
    elif mode == "missing_reason":
        selected = [
            item for item in clients
            if any(
                item.products[key] == "no" and not item.product_notes.get(key, "").strip()
                for key in PRODUCT_COLUMNS
            )
        ]
    else:
        import random
        selected = random.sample(clients, min(10, len(clients)))
    selected = selected[:100]
    await state.update_data(crm_pick=[{"identity": item.identity, "name": item.name} for item in selected])
    preview = "\n".join(f"• {html.escape(item.name)}" for item in selected[:15])
    if len(selected) > 15:
        preview += f"\n…и ещё {len(selected) - 15}"
    await render(
        callback.message,
        f"🎯 <b>Подходящих клиентов: {len(selected)}</b>\n\n{preview or 'Никого не найдено.'}",
        InlineKeyboardMarkup(inline_keyboard=(
            [[InlineKeyboardButton(text="➕ Добавить всех в мой список", callback_data="crm:pick_add")]]
            if selected else []
        ) + [nav("crm:pick")]),
    )


@router.callback_query(F.data == "crm:pick_add")
async def pick_add(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    selected = (await state.get_data()).get("crm_pick", [])
    added = sum(
        database.add_to_call_list(callback.from_user.id, item["identity"], item["name"])
        for item in selected
    )
    await callback.message.answer(f"✅ Добавлено в список: <b>{added}</b>")
    await call_list(callback, state)


@router.callback_query(F.data == "crm:report")
async def daily_report(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    await state.clear()
    events = database.daily_events(callback.from_user.id)
    if not events:
        text = "📝 <b>Итоги дня</b>\n\nСегодня через бота пока не зафиксировано ни одного действия."
        await render(callback.message, text, InlineKeyboardMarkup(inline_keyboard=[nav()]))
        await callback.answer()
        return
    counts = Counter(row["kind"] for row in events)
    grouped = {}
    for row in events:
        client = grouped.setdefault(row["client_identity"], {
            "name": row["client_name"], "note": "", "order": False,
            "status": "", "tasks": [], "product_changed": False,
        })
        if row["kind"] == "note":
            client["note"] = row["summary"]
        elif row["kind"] == "order":
            client["order"] = True
        elif row["kind"] == "status":
            client["status"] = row["summary"]
        elif row["kind"] == "task":
            client["tasks"].append(row["summary"])
        elif row["kind"] == "product":
            client["product_changed"] = True

    result_clients = sum(bool(item["note"]) for item in grouped.values())
    header = (
        f"📝 <b>Итоги за {date.today().strftime('%d.%m.%Y')}</b>\n\n"
        f"Клиентов обработано: <b>{len(grouped)}</b>\n"
        f"Зафиксировано результатов: <b>{result_clients}</b>\n"
        f"Отмечено заказов: <b>{sum(item['order'] for item in grouped.values())}</b>\n"
        f"Поставлено следующих задач: <b>{counts['task']}</b>\n\n"
        "<b>Результаты по клиентам</b>"
    )
    blocks = []
    for item in grouped.values():
        details = []
        if item["note"]:
            details.append(html.escape(item["note"]))
        if item["order"]:
            details.append("✅ Сделал заказ")
        if item["status"]:
            details.append(html.escape(item["status"]))
        if item["tasks"]:
            details.append("Следующий шаг: " + html.escape(item["tasks"][-1]))
        if not details and item["product_changed"]:
            details.append("Актуализирована информация по товарным группам")
        blocks.append(f"• <b>{html.escape(item['name'])}</b>\n" + "\n".join(details))

    chunks = []
    current = header
    for block in blocks:
        candidate = f"{current}\n\n{block}"
        if len(candidate) > 3800 and current:
            chunks.append(current)
            current = block
        else:
            current = candidate
    if current:
        chunks.append(current)
    markup = InlineKeyboardMarkup(inline_keyboard=[nav()])
    try:
        await callback.message.edit_text(chunks[0], reply_markup=markup if len(chunks) == 1 else None)
    except Exception:
        await callback.message.answer(chunks[0], reply_markup=markup if len(chunks) == 1 else None)
    for index, chunk in enumerate(chunks[1:], 1):
        await callback.message.answer(chunk, reply_markup=markup if index == len(chunks) - 1 else None)
    await callback.answer()


@router.callback_query(F.data == "crm:card")
async def back_to_card(callback: CallbackQuery, state: FSMContext) -> None:
    if not await guard(callback):
        return
    client = await current_client(state)
    if not client:
        await callback.answer("Откройте клиента через поиск повторно.", show_alert=True)
        return
    await show_client(callback.message, state, client)
    await callback.answer()
