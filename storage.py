"""SQLite persistence for MAX user preferences."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator


class UserStorage:
    def __init__(self, db_path: str | Path):
        self.db_path = str(db_path)
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as conn:
            conn.execute(
                "CREATE TABLE IF NOT EXISTS users ("
                "user_id INTEGER PRIMARY KEY, group_code TEXT NOT NULL, "
                "updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP)"
            )

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def get_group(self, user_id: int) -> str | None:
        with self._connect() as conn:
            row = conn.execute("SELECT group_code FROM users WHERE user_id = ?", (int(user_id),)).fetchone()
            return str(row["group_code"]) if row else None

    def set_group(self, user_id: int, group_code: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO users(user_id, group_code, updated_at) VALUES(?, ?, CURRENT_TIMESTAMP) "
                "ON CONFLICT(user_id) DO UPDATE SET group_code=excluded.group_code, updated_at=CURRENT_TIMESTAMP",
                (int(user_id), str(group_code)),
            )

    def clear_group(self, user_id: int) -> bool:
        """Forget one user's saved group; other users are not affected."""
        with self._connect() as conn:
            cursor = conn.execute("DELETE FROM users WHERE user_id = ?", (int(user_id),))
            return cursor.rowcount > 0

    def all_users(self) -> list[tuple[int, str]]:
        with self._connect() as conn:
            rows = conn.execute("SELECT user_id, group_code FROM users").fetchall()
        return [(int(row["user_id"]), str(row["group_code"])) for row in rows]

    def users_for_groups(self, group_codes: list[str]) -> list[tuple[int, str]]:
        if not group_codes:
            return []
        placeholders = ",".join("?" for _ in group_codes)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT user_id, group_code FROM users WHERE group_code IN ({placeholders})",
                tuple(group_codes),
            ).fetchall()
        return [(int(row["user_id"]), str(row["group_code"])) for row in rows]
