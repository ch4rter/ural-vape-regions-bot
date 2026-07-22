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


@dataclass(frozen=True)
class ClientChat:
    chat_id: int
    title: str
    chat_type: str
    is_active: bool


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
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
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
            wait_columns = {row[1] for row in connection.execute("PRAGMA table_info(wait_entries)")}
            if "source_message_id" not in wait_columns:
                connection.execute("ALTER TABLE wait_entries ADD COLUMN source_message_id INTEGER")
            if "comment" not in wait_columns:
                connection.execute("ALTER TABLE wait_entries ADD COLUMN comment TEXT NOT NULL DEFAULT ''")
            if "last_match_json" not in wait_columns:
                connection.execute("ALTER TABLE wait_entries ADD COLUMN last_match_json TEXT")

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
                "SELECT id, telegram_id, username, role FROM access_users WHERE id = ?", (access_id,)
            ).fetchone()
        return self._access_user(row) if row else None

    def list_access_users(self) -> list[AccessUser]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id, telegram_id, username, role FROM access_users ORDER BY created_at, id"
            ).fetchall()
        return [self._access_user(row) for row in rows]

    def delete_access_user(self, access_id: int) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM access_users WHERE id = ?", (access_id,))

    def authorize_user(self, telegram_id: int, username: str | None) -> bool:
        normalized_username = username.lower() if username else None
        with self._connect() as connection:
            by_id = connection.execute(
                "SELECT id FROM access_users WHERE telegram_id = ?", (telegram_id,)
            ).fetchone()
            if by_id:
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

    def upsert_client_chat(self, chat_id: int, title: str, chat_type: str, is_active: bool = True) -> None:
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO client_chats(chat_id, title, chat_type, is_active, updated_at)
                   VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
                   ON CONFLICT(chat_id) DO UPDATE SET title=excluded.title,
                       chat_type=excluded.chat_type, is_active=excluded.is_active,
                       updated_at=CURRENT_TIMESTAMP""",
                (chat_id, title.strip() or str(chat_id), chat_type, int(is_active)),
            )

    def list_client_chats(self, active_only: bool = False) -> list[ClientChat]:
        where = "WHERE is_active = 1" if active_only else ""
        with self._connect() as connection:
            rows = connection.execute(
                f"SELECT chat_id, title, chat_type, is_active FROM client_chats {where} ORDER BY title"
            ).fetchall()
        return [ClientChat(row["chat_id"], row["title"], row["chat_type"], bool(row["is_active"])) for row in rows]

    def get_client_chat(self, chat_id: int) -> ClientChat | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT chat_id, title, chat_type, is_active FROM client_chats WHERE chat_id = ?", (chat_id,)
            ).fetchone()
        return ClientChat(row["chat_id"], row["title"], row["chat_type"], bool(row["is_active"])) if row else None

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

    @staticmethod
    def _access_user(row: sqlite3.Row) -> AccessUser:
        return AccessUser(row["id"], row["telegram_id"], row["username"], row["role"])

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
                "SELECT id, product_id, name FROM sections WHERE id = ?", (section_id,)
            ).fetchone()
        return Section(row["id"], row["product_id"], row["name"]) if row else None

    def list_sections(self, product_id: int) -> list[Section]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT id, product_id, name FROM sections WHERE product_id = ? ORDER BY position, name",
                (product_id,),
            ).fetchall()
        return [Section(row["id"], row["product_id"], row["name"]) for row in rows]

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
