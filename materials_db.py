import sqlite3
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
                """
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
