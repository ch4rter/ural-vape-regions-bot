import sqlite3
import re
import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Product:
    id: int
    name: str
    is_visible: bool


@dataclass(frozen=True)
class Section:
    id: int
    product_id: int
    name: str
    is_public: bool = False


@dataclass(frozen=True)
class Material:
    id: int
    section_id: int
    kind: str
    text: str | None
    file_id: str | None
    caption: str | None
    file_name: str | None


@dataclass(frozen=True)
class AccessUser:
    id: int
    telegram_id: int | None
    username: str | None
    role: str = "user"
    sales_channel_name: str | None = None
    sales_channel_href: str | None = None


@dataclass(frozen=True)
class ClientChat:
    chat_id: int
    title: str
    chat_type: str
    is_active: bool
    tags: tuple[str, ...] = ()
    added_by_id: int | None = None
    added_by_username: str | None = None


@dataclass(frozen=True)
class WaitEntry:
    id: int
    chat_id: int
    client_title: str
    manager_id: int
    manager_name: str
    query: str
    status: str
    source_message_id: int | None
    comment: str
    last_match: dict | None


@dataclass(frozen=True)
class LeadProfile:
    user_id: int
    username: str | None
    full_name: str
    position: str
    region: str
    company: str
    outlets: str
    manager: str
    created_at: str


@dataclass(frozen=True)
class UnmatchedRegion:
    id: int
    user_id: int
    username: str | None
    raw_region: str
    normalized_region: str
    status: str
    created_at: str
    updated_at: str


class MaterialsDB:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL COLLATE NOCASE UNIQUE,
                    is_visible INTEGER NOT NULL DEFAULT 1,
                    position INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS sections (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
                    name TEXT NOT NULL COLLATE NOCASE,
                    position INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(product_id, name)
                );
                CREATE TABLE IF NOT EXISTS materials (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    section_id INTEGER NOT NULL REFERENCES sections(id) ON DELETE CASCADE,
                    kind TEXT NOT NULL CHECK(kind IN ('text', 'photo', 'document')),
                    text TEXT,
                    file_id TEXT,
                    caption TEXT,
                    file_name TEXT,
                    position INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS access_users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    telegram_id INTEGER UNIQUE,
                    username TEXT COLLATE NOCASE UNIQUE,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    CHECK(telegram_id IS NOT NULL OR username IS NOT NULL)
                );
                CREATE TABLE IF NOT EXISTS client_chats (
                    chat_id INTEGER PRIMARY KEY,
                    title TEXT NOT NULL,
                    chat_type TEXT NOT NULL,
                    is_active INTEGER NOT NULL DEFAULT 1,
                    added_by_id INTEGER,
                    added_by_username TEXT,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS client_chat_tags (
                    chat_id INTEGER NOT NULL REFERENCES client_chats(chat_id) ON DELETE CASCADE,
                    tag TEXT NOT NULL,
                    PRIMARY KEY(chat_id, tag)
                );
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS lead_profiles (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    full_name TEXT NOT NULL,
                    position TEXT NOT NULL,
                    region TEXT NOT NULL,
                    company TEXT NOT NULL,
                    outlets TEXT NOT NULL,
                    manager TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS known_telegram_users (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    full_name TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS unmatched_regions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    username TEXT,
                    raw_region TEXT NOT NULL,
                    normalized_region TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'pending'
                        CHECK(status IN ('pending', 'resolved', 'ignored')),
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(user_id, normalized_region)
                );
                CREATE INDEX IF NOT EXISTS idx_unmatched_regions_status
                    ON unmatched_regions(status, updated_at);
                CREATE TABLE IF NOT EXISTS wait_entries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    chat_id INTEGER NOT NULL,
                    client_title TEXT NOT NULL,
                    manager_id INTEGER NOT NULL,
                    manager_name TEXT NOT NULL,
                    query TEXT NOT NULL,
                    source_message_id INTEGER,
                    comment TEXT NOT NULL DEFAULT '',
                    last_match_json TEXT,
                    status TEXT NOT NULL DEFAULT 'active' CHECK(status IN ('active', 'closed')),
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    closed_at TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_wait_entries_manager
                    ON wait_entries(manager_id, status);
                CREATE TABLE IF NOT EXISTS wait_notifications (
                    wait_id INTEGER NOT NULL REFERENCES wait_entries(id) ON DELETE CASCADE,
                    item_signature TEXT NOT NULL,
                    decision TEXT NOT NULL DEFAULT 'pending'
                        CHECK(decision IN ('pending', 'confirmed', 'rejected')),
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY(wait_id, item_signature)
                );
                """
            )
            columns = {row[1] for row in connection.execute("PRAGMA table_info(access_users)")}
            if "role" not in columns:
                connection.execute("ALTER TABLE access_users ADD COLUMN role TEXT NOT NULL DEFAULT 'user'")
            if "sales_channel_name" not in columns:
                connection.execute("ALTER TABLE access_users ADD COLUMN sales_channel_name TEXT")
            if "sales_channel_href" not in columns:
                connection.execute("ALTER TABLE access_users ADD COLUMN sales_channel_href TEXT")
            wait_columns = {row[1] for row in connection.execute("PRAGMA table_info(wait_entries)")}
            if "source_message_id" not in wait_columns:
                connection.execute("ALTER TABLE wait_entries ADD COLUMN source_message_id INTEGER")
            if "comment" not in wait_columns:
                connection.execute("ALTER TABLE wait_entries ADD COLUMN comment TEXT NOT NULL DEFAULT ''")
            if "last_match_json" not in wait_columns:
                connection.execute("ALTER TABLE wait_entries ADD COLUMN last_match_json TEXT")
            section_columns = {row[1] for row in connection.execute("PRAGMA table_info(sections)")}
            if "is_public" not in section_columns:
                connection.execute("ALTER TABLE sections ADD COLUMN is_public INTEGER NOT NULL DEFAULT 0")
            chat_columns = {row[1] for row in connection.execute("PRAGMA table_info(client_chats)")}
            if "added_by_id" not in chat_columns:
                connection.execute("ALTER TABLE client_chats ADD COLUMN added_by_id INTEGER")
            if "added_by_username" not in chat_columns:
                connection.execute("ALTER TABLE client_chats ADD COLUMN added_by_username TEXT")

    def add_access_user(self, value: str) -> AccessUser:
        value = value.strip()
        telegram_id = None
        username = None
        if not value.startswith("@") and value.isdigit():
            telegram_id = int(value)
            if telegram_id <= 0:
                raise ValueError("Telegram ID должен быть положительным числом.")
        else:
            username = value.lstrip("@").lower()
            if not re.fullmatch(r"[a-zA-Z0-9_]{5,32}", username):
                raise ValueError("Username должен содержать 5–32 латинских символа, цифры или _. ")
        with self._connect() as connection:
            cursor = connection.execute(
                "INSERT INTO access_users(telegram_id, username) VALUES (?, ?)",
                (telegram_id, username),
            )
            access_id = cursor.lastrowid
            connection.commit()
        return self.get_access_user(access_id)

    def get_access_user(self, access_id: int) -> AccessUser | None:
        with self._connect() as connection:
            row = connection.execute(
                """SELECT id, telegram_id, username, role,
                          sales_channel_name, sales_channel_href
                   FROM access_users WHERE id = ?""", (access_id,)
            ).fetchone()
        return self._access_user(row) if row else None

    def list_access_users(self) -> list[AccessUser]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT id, telegram_id, username, role,
                          sales_channel_name, sales_channel_href
                   FROM access_users ORDER BY created_at, id"""
            ).fetchall()
        return [self._access_user(row) for row in rows]

    def delete_access_user(self, access_id: int) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM access_users WHERE id = ?", (access_id,))

    def authorize_user(self, telegram_id: int, username: str | None) -> bool:
        normalized_username = username.lower() if username else None
        with self._connect() as connection:
            by_id = connection.execute(
                "SELECT id, username FROM access_users WHERE telegram_id = ?", (telegram_id,)
            ).fetchone()
            if by_id:
                if normalized_username and not by_id["username"]:
                    try:
                        connection.execute(
                            "UPDATE access_users SET username = ? WHERE id = ? AND username IS NULL",
                            (normalized_username, by_id["id"]),
                        )
                        connection.commit()
                    except sqlite3.IntegrityError:
                        pass
                return True
            if not normalized_username:
                return False
            by_username = connection.execute(
                "SELECT id, telegram_id FROM access_users WHERE username = ? COLLATE NOCASE",
                (normalized_username,),
            ).fetchone()
            if not by_username:
                return False
            if by_username["telegram_id"] is not None:
                return by_username["telegram_id"] == telegram_id
            try:
                connection.execute(
                    "UPDATE access_users SET telegram_id = ? WHERE id = ? AND telegram_id IS NULL",
                    (telegram_id, by_username["id"]),
                )
                connection.commit()
            except sqlite3.IntegrityError:
                return False
            return True

    def user_role(self, telegram_id: int, username: str | None = None) -> str | None:
        if not self.authorize_user(telegram_id, username):
            return None
        with self._connect() as connection:
            row = connection.execute(
                "SELECT role FROM access_users WHERE telegram_id = ?", (telegram_id,)
            ).fetchone()
        return row["role"] if row else None

    def set_access_role(self, access_id: int, role: str) -> None:
        if role not in {"user", "junior_admin"}:
            raise ValueError("Неизвестная роль пользователя.")
        with self._connect() as connection:
            connection.execute("UPDATE access_users SET role = ? WHERE id = ?", (role, access_id))

    def set_access_sales_channel(
        self, access_id: int, name: str | None, href: str | None
    ) -> None:
        if (name is None) != (href is None):
            raise ValueError("Название и ссылка канала продаж должны задаваться вместе.")
        with self._connect() as connection:
            connection.execute(
                """UPDATE access_users
                   SET sales_channel_name = ?, sales_channel_href = ?
                   WHERE id = ?""",
                (name, href, access_id),
            )

    def user_sales_channel(
        self, telegram_id: int, username: str | None = None
    ) -> tuple[str, str] | None:
        if not self.authorize_user(telegram_id, username):
            return None
        with self._connect() as connection:
            row = connection.execute(
                """SELECT sales_channel_name, sales_channel_href
                   FROM access_users WHERE telegram_id = ?""",
                (telegram_id,),
            ).fetchone()
        if not row or not row["sales_channel_name"] or not row["sales_channel_href"]:
            return None
        return row["sales_channel_name"], row["sales_channel_href"]

    def upsert_client_chat(
        self, chat_id: int, title: str, chat_type: str, is_active: bool = True,
        added_by_id: int | None = None, added_by_username: str | None = None,
    ) -> None:
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO client_chats(
                       chat_id, title, chat_type, is_active, added_by_id, added_by_username, updated_at
                   ) VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                   ON CONFLICT(chat_id) DO UPDATE SET title=excluded.title,
                       chat_type=excluded.chat_type, is_active=excluded.is_active,
                       added_by_id=COALESCE(excluded.added_by_id, client_chats.added_by_id),
                       added_by_username=COALESCE(
                           excluded.added_by_username, client_chats.added_by_username
                       ),
                       updated_at=CURRENT_TIMESTAMP""",
                (
                    chat_id, title.strip() or str(chat_id), chat_type, int(is_active),
                    added_by_id, added_by_username,
                ),
            )

    def list_client_chats(
        self, active_only: bool = False, tags: tuple[str, ...] | list[str] | None = None,
        untagged_only: bool = False,
    ) -> list[ClientChat]:
        conditions, values = [], []
        if active_only:
            conditions.append("c.is_active = 1")
        normalized_tags = tuple(dict.fromkeys(tag.strip().lower() for tag in (tags or ()) if tag.strip()))
        if normalized_tags:
            placeholders = ", ".join("?" for _ in normalized_tags)
            conditions.append(
                f"EXISTS(SELECT 1 FROM client_chat_tags f WHERE f.chat_id = c.chat_id "
                f"AND f.tag IN ({placeholders}))"
            )
            values.extend(normalized_tags)
        if untagged_only:
            conditions.append("NOT EXISTS(SELECT 1 FROM client_chat_tags u WHERE u.chat_id = c.chat_id)")
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        with self._connect() as connection:
            rows = connection.execute(
                f"""SELECT c.chat_id, c.title, c.chat_type, c.is_active,
                           c.added_by_id, c.added_by_username,
                           GROUP_CONCAT(t.tag, ',') AS tags
                    FROM client_chats c
                    LEFT JOIN client_chat_tags t ON t.chat_id = c.chat_id
                    {where}
                    GROUP BY c.chat_id
                    ORDER BY c.title""",
                values,
            ).fetchall()
        return [self._client_chat(row) for row in rows]

    def get_client_chat(self, chat_id: int) -> ClientChat | None:
        with self._connect() as connection:
            row = connection.execute(
                """SELECT c.chat_id, c.title, c.chat_type, c.is_active,
                          c.added_by_id, c.added_by_username,
                          GROUP_CONCAT(t.tag, ',') AS tags
                   FROM client_chats c
                   LEFT JOIN client_chat_tags t ON t.chat_id = c.chat_id
                   WHERE c.chat_id = ?
                   GROUP BY c.chat_id""",
                (chat_id,),
            ).fetchone()
        return self._client_chat(row) if row else None

    def set_client_chat_tags(self, chat_id: int, tags: list[str] | tuple[str, ...]) -> None:
        normalized = tuple(dict.fromkeys(tag.strip().lower() for tag in tags if tag.strip()))
        with self._connect() as connection:
            if not connection.execute(
                "SELECT 1 FROM client_chats WHERE chat_id = ?", (chat_id,)
            ).fetchone():
                raise ValueError("Клиентский чат не найден.")
            connection.execute("DELETE FROM client_chat_tags WHERE chat_id = ?", (chat_id,))
            connection.executemany(
                "INSERT INTO client_chat_tags(chat_id, tag) VALUES (?, ?)",
                [(chat_id, tag) for tag in normalized],
            )

    def toggle_client_chat_tag(self, chat_id: int, tag: str) -> tuple[str, ...]:
        normalized = tag.strip().lower()
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT 1 FROM client_chat_tags WHERE chat_id = ? AND tag = ?",
                (chat_id, normalized),
            ).fetchone()
            if existing:
                connection.execute(
                    "DELETE FROM client_chat_tags WHERE chat_id = ? AND tag = ?",
                    (chat_id, normalized),
                )
            else:
                connection.execute(
                    "INSERT INTO client_chat_tags(chat_id, tag) VALUES (?, ?)",
                    (chat_id, normalized),
                )
        chat = self.get_client_chat(chat_id)
        return chat.tags if chat else ()

    @staticmethod
    def _client_chat(row: sqlite3.Row) -> ClientChat:
        tags = tuple(sorted(filter(None, (row["tags"] or "").split(","))))
        return ClientChat(
            row["chat_id"], row["title"], row["chat_type"], bool(row["is_active"]),
            tags, row["added_by_id"], row["added_by_username"],
        )

    def add_wait_entry(
        self, chat_id: int, client_title: str, manager_id: int, manager_name: str, query: str,
        source_message_id: int | None = None, comment: str = "",
    ) -> WaitEntry:
        query = query.strip()
        if len(query) < 2:
            raise ValueError("Название товара слишком короткое.")
        with self._connect() as connection:
            active_rows = connection.execute(
                """SELECT id, query, client_title FROM wait_entries
                   WHERE chat_id = ? AND manager_id = ? AND status = 'active'""",
                (chat_id, manager_id),
            ).fetchall()
            normalized_query = " ".join(query.casefold().replace("ё", "е").split())
            normalized_client = " ".join(client_title.casefold().replace("ё", "е").split())
            existing = next(
                (
                    row for row in active_rows
                    if " ".join(row["query"].casefold().replace("ё", "е").split()) == normalized_query
                    and (
                        chat_id != 0
                        or " ".join(row["client_title"].casefold().replace("ё", "е").split())
                        == normalized_client
                    )
                ),
                None,
            )
            if existing:
                wait_id = existing["id"]
                connection.execute(
                    """UPDATE wait_entries SET client_title = ?, manager_name = ?,
                           source_message_id = COALESCE(?, source_message_id),
                           comment = CASE WHEN ? <> '' THEN ? ELSE comment END
                       WHERE id = ?""",
                    (
                        client_title.strip() or str(chat_id), manager_name, source_message_id,
                        comment.strip(), comment.strip(), wait_id,
                    ),
                )
            else:
                cursor = connection.execute(
                    """INSERT INTO wait_entries(
                           chat_id, client_title, manager_id, manager_name, query, source_message_id, comment
                       ) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (
                        chat_id, client_title.strip() or str(chat_id), manager_id, manager_name,
                        query, source_message_id, comment.strip(),
                    ),
                )
                wait_id = cursor.lastrowid
            connection.commit()
        return self.get_wait_entry(wait_id)

    def get_wait_entry(self, wait_id: int) -> WaitEntry | None:
        with self._connect() as connection:
            row = connection.execute(
                """SELECT id, chat_id, client_title, manager_id, manager_name, query, status,
                          source_message_id, comment, last_match_json
                   FROM wait_entries WHERE id = ?""",
                (wait_id,),
            ).fetchone()
        return self._wait_entry(row) if row else None

    def list_wait_entries(self, manager_id: int | None = None, active_only: bool = True) -> list[WaitEntry]:
        conditions, values = [], []
        if manager_id is not None:
            conditions.append("manager_id = ?")
            values.append(manager_id)
        if active_only:
            conditions.append("status = 'active'")
        where = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        with self._connect() as connection:
            rows = connection.execute(
                f"""SELECT id, chat_id, client_title, manager_id, manager_name, query, status,
                           source_message_id, comment, last_match_json
                    FROM wait_entries {where} ORDER BY created_at DESC, id DESC""",
                values,
            ).fetchall()
        return [self._wait_entry(row) for row in rows]

    def update_wait_entry(
        self, wait_id: int, manager_id: int, *, query: str | None = None, comment: str | None = None
    ) -> WaitEntry | None:
        fields, values = [], []
        query_changed = query is not None
        if query is not None:
            query = query.strip()
            if len(query) < 2:
                raise ValueError("Название товара слишком короткое.")
            fields.append("query = ?")
            values.append(query)
            fields.append("last_match_json = NULL")
        if comment is not None:
            fields.append("comment = ?")
            values.append(comment.strip())
        if not fields:
            return self.get_wait_entry(wait_id)
        values.extend([wait_id, manager_id])
        with self._connect() as connection:
            connection.execute(
                f"UPDATE wait_entries SET {', '.join(fields)} WHERE id = ? AND manager_id = ? AND status = 'active'",
                values,
            )
            if query_changed:
                connection.execute("DELETE FROM wait_notifications WHERE wait_id = ?", (wait_id,))
        return self.get_wait_entry(wait_id)

    def set_wait_last_match(self, wait_id: int, payload: dict) -> None:
        with self._connect() as connection:
            connection.execute(
                "UPDATE wait_entries SET last_match_json = ? WHERE id = ?",
                (json.dumps(payload, ensure_ascii=False), wait_id),
            )

    def close_wait_entry(self, wait_id: int, manager_id: int | None = None) -> bool:
        condition = "id = ?" if manager_id is None else "id = ? AND manager_id = ?"
        values = (wait_id,) if manager_id is None else (wait_id, manager_id)
        with self._connect() as connection:
            cursor = connection.execute(
                f"UPDATE wait_entries SET status = 'closed', closed_at = CURRENT_TIMESTAMP WHERE {condition}",
                values,
            )
        return cursor.rowcount > 0

    def wait_match_seen(self, wait_id: int, item_signature: str) -> bool:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT 1 FROM wait_notifications WHERE wait_id = ? AND item_signature = ?",
                (wait_id, item_signature),
            ).fetchone()
        return bool(row)

    def record_wait_match(self, wait_id: int, item_signature: str) -> None:
        with self._connect() as connection:
            connection.execute(
                """INSERT OR IGNORE INTO wait_notifications(wait_id, item_signature)
                   VALUES (?, ?)""",
                (wait_id, item_signature),
            )

    def reject_wait_match(self, wait_id: int, item_signature: str, manager_id: int) -> bool:
        with self._connect() as connection:
            cursor = connection.execute(
                """UPDATE wait_notifications SET decision = 'rejected'
                   WHERE wait_id = ? AND item_signature = ?
                     AND EXISTS(SELECT 1 FROM wait_entries WHERE id = ? AND manager_id = ?)""",
                (wait_id, item_signature, wait_id, manager_id),
            )
        return cursor.rowcount > 0

    def set_setting(self, key: str, value: str) -> None:
        with self._connect() as connection:
            connection.execute(
                "INSERT INTO settings(key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )

    def get_setting(self, key: str) -> str | None:
        with self._connect() as connection:
            row = connection.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else None

    def save_lead_profile(
        self, user_id: int, username: str | None, full_name: str, position: str,
        region: str, company: str, outlets: str, manager: str,
    ) -> LeadProfile:
        values = [full_name, position, region, company, outlets, manager]
        if any(not value.strip() for value in values):
            raise ValueError("Все поля анкеты должны быть заполнены.")
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO lead_profiles(
                       user_id, username, full_name, position, region, company, outlets, manager
                   ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(user_id) DO NOTHING""",
                (
                    user_id, username.lower() if username else None, full_name.strip(),
                    position.strip(), region.strip(), company.strip(), outlets.strip(), manager.strip(),
                ),
            )
        profile = self.get_lead_profile(user_id)
        if not profile:
            raise RuntimeError("Не удалось сохранить анкету.")
        return profile

    def get_lead_profile(self, user_id: int) -> LeadProfile | None:
        with self._connect() as connection:
            row = connection.execute(
                """SELECT user_id, username, full_name, position, region, company,
                          outlets, manager, created_at
                   FROM lead_profiles WHERE user_id = ?""",
                (user_id,),
            ).fetchone()
        return self._lead_profile(row) if row else None

    def list_lead_profiles(self) -> list[LeadProfile]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT user_id, username, full_name, position, region, company,
                          outlets, manager, created_at
                   FROM lead_profiles ORDER BY created_at DESC, user_id DESC"""
            ).fetchall()
        return [self._lead_profile(row) for row in rows]

    def save_unmatched_region(
        self, user_id: int, username: str | None, raw_region: str, normalized_region: str
    ) -> UnmatchedRegion:
        raw_region = raw_region.strip()
        normalized_region = normalized_region.strip()
        if not raw_region or not normalized_region:
            raise ValueError("Регион не может быть пустым.")
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO unmatched_regions(
                       user_id, username, raw_region, normalized_region
                   ) VALUES (?, ?, ?, ?)
                   ON CONFLICT(user_id, normalized_region) DO UPDATE SET
                       username=excluded.username,
                       raw_region=excluded.raw_region,
                       status='pending',
                       updated_at=CURRENT_TIMESTAMP""",
                (user_id, username.lower() if username else None, raw_region, normalized_region),
            )
            row = connection.execute(
                """SELECT id, user_id, username, raw_region, normalized_region,
                          status, created_at, updated_at
                   FROM unmatched_regions
                   WHERE user_id = ? AND normalized_region = ?""",
                (user_id, normalized_region),
            ).fetchone()
        return self._unmatched_region(row)

    def list_unmatched_regions(self, status: str = "pending") -> list[UnmatchedRegion]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT id, user_id, username, raw_region, normalized_region,
                          status, created_at, updated_at
                   FROM unmatched_regions WHERE status = ?
                   ORDER BY updated_at DESC, id DESC""",
                (status,),
            ).fetchall()
        return [self._unmatched_region(row) for row in rows]

    def get_unmatched_region(self, region_id: int) -> UnmatchedRegion | None:
        with self._connect() as connection:
            row = connection.execute(
                """SELECT id, user_id, username, raw_region, normalized_region,
                          status, created_at, updated_at
                   FROM unmatched_regions WHERE id = ?""",
                (region_id,),
            ).fetchone()
        return self._unmatched_region(row) if row else None

    def set_unmatched_region_status(self, region_id: int, status: str) -> bool:
        if status not in {"pending", "resolved", "ignored"}:
            raise ValueError("Некорректный статус.")
        with self._connect() as connection:
            cursor = connection.execute(
                """UPDATE unmatched_regions SET status = ?, updated_at=CURRENT_TIMESTAMP
                   WHERE id = ?""",
                (status, region_id),
            )
        return cursor.rowcount > 0

    def resolve_unmatched_regions(self, normalized_region: str) -> int:
        with self._connect() as connection:
            cursor = connection.execute(
                """UPDATE unmatched_regions SET status = 'resolved', updated_at=CURRENT_TIMESTAMP
                   WHERE normalized_region = ? AND status = 'pending'""",
                (normalized_region.strip(),),
            )
        return cursor.rowcount

    def remember_telegram_user(
        self, user_id: int, username: str | None, full_name: str = ""
    ) -> None:
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO known_telegram_users(user_id, username, full_name, updated_at)
                   VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                   ON CONFLICT(user_id) DO UPDATE SET
                       username=COALESCE(excluded.username, known_telegram_users.username),
                       full_name=CASE
                           WHEN excluded.full_name <> '' THEN excluded.full_name
                           ELSE known_telegram_users.full_name
                       END,
                       updated_at=CURRENT_TIMESTAMP""",
                (user_id, username.lower() if username else None, full_name.strip()),
            )

    def telegram_user_id_by_username(self, username: str) -> int | None:
        normalized = username.strip().lstrip("@").lower()
        with self._connect() as connection:
            row = connection.execute(
                """SELECT user_id FROM known_telegram_users
                   WHERE username = ? COLLATE NOCASE
                   ORDER BY updated_at DESC LIMIT 1""",
                (normalized,),
            ).fetchone()
        return row["user_id"] if row else None

    @staticmethod
    def _lead_profile(row: sqlite3.Row) -> LeadProfile:
        return LeadProfile(
            row["user_id"], row["username"], row["full_name"], row["position"],
            row["region"], row["company"], row["outlets"], row["manager"], row["created_at"],
        )

    @staticmethod
    def _unmatched_region(row: sqlite3.Row) -> UnmatchedRegion:
        return UnmatchedRegion(
            row["id"], row["user_id"], row["username"], row["raw_region"],
            row["normalized_region"], row["status"], row["created_at"], row["updated_at"],
        )

    @staticmethod
    def _access_user(row: sqlite3.Row) -> AccessUser:
        return AccessUser(
            row["id"], row["telegram_id"], row["username"], row["role"],
            row["sales_channel_name"], row["sales_channel_href"],
        )

    @staticmethod
    def _wait_entry(row: sqlite3.Row) -> WaitEntry:
        return WaitEntry(
            row["id"], row["chat_id"], row["client_title"], row["manager_id"],
            row["manager_name"], row["query"], row["status"], row["source_message_id"],
            row["comment"], json.loads(row["last_match_json"]) if row["last_match_json"] else None,
        )

    def backup_to(self, destination: Path) -> None:
        """Create a consistent SQLite backup, including pending WAL changes."""
        destination.parent.mkdir(parents=True, exist_ok=True)
        source = self._connect()
        target = sqlite3.connect(destination)
        try:
            source.backup(target)
        finally:
            target.close()
            source.close()

    def add_product(self, name: str) -> Product:
        with self._connect() as connection:
            position = connection.execute("SELECT COALESCE(MAX(position), -1) + 1 FROM products").fetchone()[0]
            cursor = connection.execute(
                "INSERT INTO products(name, position) VALUES (?, ?)", (name.strip(), position)
            )
            product_id = cursor.lastrowid
            connection.commit()
        return self.get_product(product_id)

    def get_product(self, product_id: int) -> Product | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT id, name, is_visible FROM products WHERE id = ?", (product_id,)
            ).fetchone()
        return Product(row["id"], row["name"], bool(row["is_visible"])) if row else None

    def list_products(self, visible_only: bool = False) -> list[Product]:
        where = "WHERE is_visible = 1" if visible_only else ""
        with self._connect() as connection:
            rows = connection.execute(
                f"SELECT id, name, is_visible FROM products {where} ORDER BY position, name"
            ).fetchall()
        return [Product(row["id"], row["name"], bool(row["is_visible"])) for row in rows]

    def rename_product(self, product_id: int, name: str) -> None:
        with self._connect() as connection:
            connection.execute("UPDATE products SET name = ? WHERE id = ?", (name.strip(), product_id))

    def toggle_product(self, product_id: int) -> None:
        with self._connect() as connection:
            connection.execute(
                "UPDATE products SET is_visible = CASE is_visible WHEN 1 THEN 0 ELSE 1 END WHERE id = ?",
                (product_id,),
            )

    def delete_product(self, product_id: int) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM products WHERE id = ?", (product_id,))

    def add_section(self, product_id: int, name: str) -> Section:
        with self._connect() as connection:
            position = connection.execute(
                "SELECT COALESCE(MAX(position), -1) + 1 FROM sections WHERE product_id = ?",
                (product_id,),
            ).fetchone()[0]
            cursor = connection.execute(
                "INSERT INTO sections(product_id, name, position) VALUES (?, ?, ?)",
                (product_id, name.strip(), position),
            )
            section_id = cursor.lastrowid
            connection.commit()
        return self.get_section(section_id)

    def get_section(self, section_id: int) -> Section | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT id, product_id, name, is_public FROM sections WHERE id = ?", (section_id,)
            ).fetchone()
        return Section(row["id"], row["product_id"], row["name"], bool(row["is_public"])) if row else None

    def list_sections(self, product_id: int, public_only: bool = False) -> list[Section]:
        public_filter = "AND is_public = 1" if public_only else ""
        with self._connect() as connection:
            rows = connection.execute(
                f"""SELECT id, product_id, name, is_public FROM sections
                    WHERE product_id = ? {public_filter} ORDER BY position, name""",
                (product_id,),
            ).fetchall()
        return [
            Section(row["id"], row["product_id"], row["name"], bool(row["is_public"]))
            for row in rows
        ]

    def list_public_products(self) -> list[Product]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT DISTINCT p.id, p.name, p.is_visible, p.position
                   FROM products p
                   JOIN sections s ON s.product_id = p.id
                   WHERE p.is_visible = 1 AND s.is_public = 1
                   ORDER BY p.position, p.name"""
            ).fetchall()
        return [Product(row["id"], row["name"], bool(row["is_visible"])) for row in rows]

    def toggle_section_public(self, section_id: int) -> bool:
        with self._connect() as connection:
            connection.execute(
                "UPDATE sections SET is_public = CASE is_public WHEN 1 THEN 0 ELSE 1 END WHERE id = ?",
                (section_id,),
            )
            row = connection.execute(
                "SELECT is_public FROM sections WHERE id = ?", (section_id,)
            ).fetchone()
        if not row:
            raise ValueError("Раздел не найден.")
        return bool(row["is_public"])

    def rename_section(self, section_id: int, name: str) -> None:
        with self._connect() as connection:
            connection.execute("UPDATE sections SET name = ? WHERE id = ?", (name.strip(), section_id))

    def delete_section(self, section_id: int) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM sections WHERE id = ?", (section_id,))

    def add_material(
        self,
        section_id: int,
        kind: str,
        *,
        text: str | None = None,
        file_id: str | None = None,
        caption: str | None = None,
        file_name: str | None = None,
    ) -> Material:
        with self._connect() as connection:
            position = connection.execute(
                "SELECT COALESCE(MAX(position), -1) + 1 FROM materials WHERE section_id = ?",
                (section_id,),
            ).fetchone()[0]
            cursor = connection.execute(
                """INSERT INTO materials(section_id, kind, text, file_id, caption, file_name, position)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (section_id, kind, text, file_id, caption, file_name, position),
            )
            material_id = cursor.lastrowid
            connection.commit()
        return self.get_material(material_id)

    def get_material(self, material_id: int) -> Material | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT id, section_id, kind, text, file_id, caption, file_name FROM materials WHERE id = ?",
                (material_id,),
            ).fetchone()
        return self._material(row) if row else None

    def list_materials(self, section_id: int) -> list[Material]:
        with self._connect() as connection:
            rows = connection.execute(
                """SELECT id, section_id, kind, text, file_id, caption, file_name
                   FROM materials WHERE section_id = ? ORDER BY position, id""",
                (section_id,),
            ).fetchall()
        return [self._material(row) for row in rows]

    def delete_material(self, material_id: int) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM materials WHERE id = ?", (material_id,))

    @staticmethod
    def _material(row: sqlite3.Row) -> Material:
        return Material(
            row["id"], row["section_id"], row["kind"], row["text"], row["file_id"],
            row["caption"], row["file_name"]
        )
