"""SQLite persistence for Services Monitor snapshots."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterable
from datetime import datetime, timezone
from pathlib import Path

from models import ServiceInfo, Snapshot


SCHEMA = """
CREATE TABLE IF NOT EXISTS snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    captured_at TEXT NOT NULL,
    label TEXT
);
CREATE TABLE IF NOT EXISTS services (
    snapshot_id INTEGER NOT NULL REFERENCES snapshots(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    display_name TEXT NOT NULL,
    state TEXT NOT NULL,
    start_type TEXT NOT NULL,
    pid INTEGER NOT NULL,
    PRIMARY KEY (snapshot_id, name)
);
CREATE INDEX IF NOT EXISTS idx_services_snapshot_pid
    ON services(snapshot_id, pid);
"""


class SnapshotStore:
    def __init__(self, path: str | Path = "services.db") -> None:
        self.path = Path(path)
        self.connection = sqlite3.connect(self.path)
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.executescript(SCHEMA)

    def __enter__(self) -> "SnapshotStore":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def close(self) -> None:
        self.connection.close()

    def save_snapshot(
        self,
        services: Iterable[ServiceInfo],
        label: str | None = None,
        captured_at: str | None = None,
    ) -> Snapshot:
        timestamp = captured_at or datetime.now(timezone.utc).isoformat(timespec="seconds")
        with self.connection:
            cursor = self.connection.execute(
                "INSERT INTO snapshots(captured_at, label) VALUES (?, ?)",
                (timestamp, label),
            )
            snapshot_id = int(cursor.lastrowid)
            self.connection.executemany(
                """INSERT INTO services
                   (snapshot_id, name, display_name, state, start_type, pid)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (
                    (
                        snapshot_id,
                        item.name,
                        item.display_name,
                        item.state,
                        item.start_type,
                        item.pid,
                    )
                    for item in services
                ),
            )
        return Snapshot(snapshot_id, timestamp, label)

    def list_snapshots(self) -> list[Snapshot]:
        rows = self.connection.execute(
            "SELECT id, captured_at, label FROM snapshots ORDER BY id"
        ).fetchall()
        return [Snapshot(int(row[0]), row[1], row[2]) for row in rows]

    def get_services(self, snapshot_id: int) -> list[ServiceInfo]:
        rows = self.connection.execute(
            """SELECT name, display_name, state, start_type, pid
               FROM services WHERE snapshot_id = ? ORDER BY name COLLATE NOCASE""",
            (snapshot_id,),
        ).fetchall()
        if not rows and not self.connection.execute(
            "SELECT 1 FROM snapshots WHERE id = ?", (snapshot_id,)
        ).fetchone():
            raise ValueError(f"Snapshot {snapshot_id} does not exist")
        return [ServiceInfo(row[0], row[1], row[2], row[3], int(row[4])) for row in rows]

    def latest_ids(self, count: int = 2) -> list[int]:
        rows = self.connection.execute(
            "SELECT id FROM snapshots ORDER BY id DESC LIMIT ?", (count,)
        ).fetchall()
        return [int(row[0]) for row in reversed(rows)]
