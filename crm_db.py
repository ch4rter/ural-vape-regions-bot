import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path


@dataclass(frozen=True)
class CRMTask:
    id: int
    owner_id: int
    client_identity: str
    client_name: str
    due_date: str
    text: str
    status: str
    created_at: str
    completed_at: str | None


class CRMDatabase:
    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self):
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        return connection

    def _initialize(self):
        with self._connect() as connection:
            connection.executescript("""
                CREATE TABLE IF NOT EXISTS crm_tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    owner_id INTEGER NOT NULL,
                    client_identity TEXT NOT NULL,
                    client_name TEXT NOT NULL,
                    due_date TEXT NOT NULL,
                    text TEXT NOT NULL,
                    status TEXT NOT NULL DEFAULT 'active' CHECK(status IN ('active', 'done')),
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    completed_at TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_crm_tasks_owner_due
                    ON crm_tasks(owner_id, status, due_date);
                CREATE TABLE IF NOT EXISTS crm_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    owner_id INTEGER NOT NULL,
                    client_identity TEXT NOT NULL,
                    client_name TEXT NOT NULL,
                    kind TEXT NOT NULL,
                    summary TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                CREATE INDEX IF NOT EXISTS idx_crm_events_owner_created
                    ON crm_events(owner_id, created_at);
                CREATE TABLE IF NOT EXISTS crm_call_list (
                    owner_id INTEGER NOT NULL,
                    client_identity TEXT NOT NULL,
                    client_name TEXT NOT NULL,
                    position INTEGER NOT NULL,
                    added_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY(owner_id, client_identity)
                );
            """)

    @staticmethod
    def _task(row) -> CRMTask:
        return CRMTask(row["id"], row["owner_id"], row["client_identity"], row["client_name"],
                       row["due_date"], row["text"], row["status"], row["created_at"], row["completed_at"])

    def add_task(self, owner_id: int, identity: str, name: str, due_date: date, text: str) -> CRMTask:
        with self._connect() as connection:
            cursor = connection.execute(
                "INSERT INTO crm_tasks(owner_id, client_identity, client_name, due_date, text) VALUES (?, ?, ?, ?, ?)",
                (owner_id, identity, name, due_date.isoformat(), text.strip()),
            )
            task_id = cursor.lastrowid
            row = connection.execute("SELECT * FROM crm_tasks WHERE id = ?", (task_id,)).fetchone()
        return self._task(row)

    def tasks(self, owner_id: int, include_future: bool = False) -> list[CRMTask]:
        condition = "" if include_future else "AND due_date <= ?"
        params = (owner_id,) if include_future else (owner_id, date.today().isoformat())
        with self._connect() as connection:
            rows = connection.execute(
                f"SELECT * FROM crm_tasks WHERE owner_id = ? AND status = 'active' {condition} ORDER BY due_date, id",
                params,
            ).fetchall()
        return [self._task(row) for row in rows]

    def complete_task(self, owner_id: int, task_id: int) -> bool:
        with self._connect() as connection:
            cursor = connection.execute(
                "UPDATE crm_tasks SET status='done', completed_at=CURRENT_TIMESTAMP WHERE id=? AND owner_id=? AND status='active'",
                (task_id, owner_id),
            )
        return bool(cursor.rowcount)

    def add_event(self, owner_id: int, identity: str, name: str, kind: str, summary: str) -> None:
        with self._connect() as connection:
            connection.execute(
                "INSERT INTO crm_events(owner_id, client_identity, client_name, kind, summary) VALUES (?, ?, ?, ?, ?)",
                (owner_id, identity, name, kind, summary.strip()),
            )

    def daily_events(self, owner_id: int, day: date | None = None):
        day = day or date.today()
        with self._connect() as connection:
            return connection.execute(
                "SELECT * FROM crm_events WHERE owner_id=? AND date(created_at, 'localtime')=? ORDER BY id",
                (owner_id, day.isoformat()),
            ).fetchall()

    def add_to_call_list(self, owner_id: int, identity: str, name: str) -> bool:
        with self._connect() as connection:
            position = connection.execute(
                "SELECT COALESCE(MAX(position), 0) + 1 FROM crm_call_list WHERE owner_id=?", (owner_id,)
            ).fetchone()[0]
            cursor = connection.execute(
                "INSERT OR IGNORE INTO crm_call_list(owner_id, client_identity, client_name, position) VALUES (?, ?, ?, ?)",
                (owner_id, identity, name, position),
            )
        return bool(cursor.rowcount)

    def call_list(self, owner_id: int):
        with self._connect() as connection:
            return connection.execute(
                "SELECT * FROM crm_call_list WHERE owner_id=? ORDER BY position", (owner_id,)
            ).fetchall()

    def remove_from_call_list(self, owner_id: int, identity: str) -> None:
        with self._connect() as connection:
            connection.execute(
                "DELETE FROM crm_call_list WHERE owner_id=? AND client_identity=?", (owner_id, identity)
            )

    def clear_call_list(self, owner_id: int) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM crm_call_list WHERE owner_id=?", (owner_id,))

    def backup_to(self, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        source = self._connect()
        target = sqlite3.connect(destination)
        try:
            source.backup(target)
        finally:
            target.close()
            source.close()
