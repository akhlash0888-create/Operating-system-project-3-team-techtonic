from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from analysis import diff_services, group_by_pid
from db import SnapshotStore
from models import ServiceInfo


def service(
    name: str,
    state: str = "RUNNING",
    start_type: str = "AUTOMATIC",
    pid: int = 100,
) -> ServiceInfo:
    return ServiceInfo(name, f"{name} Display", state, start_type, pid)


class AnalysisTests(unittest.TestCase):
    def test_diff_detects_state_start_type_pid_add_and_remove(self) -> None:
        before = [service("Alpha"), service("Removed", pid=200)]
        after = [service("Alpha", "STOPPED", "MANUAL", 0), service("Added", pid=300)]
        changes = diff_services(before, after)
        self.assertEqual([item.name for item in changes], ["Added", "Alpha", "Removed"])
        alpha = changes[1]
        self.assertEqual((alpha.old_state, alpha.new_state), ("RUNNING", "STOPPED"))
        self.assertEqual((alpha.old_start_type, alpha.new_start_type), ("AUTOMATIC", "MANUAL"))
        self.assertEqual((alpha.old_pid, alpha.new_pid), (100, 0))

    def test_group_by_pid_excludes_stopped_and_orders_largest_first(self) -> None:
        items = [
            service("A", pid=10),
            service("B", pid=10),
            service("C", pid=10),
            service("D", pid=20),
            service("E", pid=20),
            service("Stopped", state="STOPPED", pid=0),
        ]
        groups = group_by_pid(items)
        self.assertEqual([(pid, len(group)) for pid, group in groups], [(10, 3), (20, 2)])


class DatabaseTests(unittest.TestCase):
    def test_snapshot_round_trip_and_latest_ids(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "test.db"
            with SnapshotStore(path) as store:
                first = store.save_snapshot([service("Alpha")], "before")
                second = store.save_snapshot([service("Alpha", "STOPPED", pid=0)], "after")
                self.assertEqual(store.latest_ids(), [first.id, second.id])
                self.assertEqual(store.get_services(first.id), [service("Alpha")])
                self.assertEqual([item.label for item in store.list_snapshots()], ["before", "after"])

    def test_missing_snapshot_raises(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            with SnapshotStore(Path(temp_dir) / "test.db") as store:
                with self.assertRaises(ValueError):
                    store.get_services(999)


if __name__ == "__main__":
    unittest.main()
