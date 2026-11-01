"""Command-line interface for the Windows Services Monitor project."""

from __future__ import annotations

import argparse
import sys
import time
from collections.abc import Iterable, Sequence

from analysis import diff_services, group_by_pid
from db import SnapshotStore
from models import ServiceChange, ServiceInfo
from scm import enumerate_services


def print_table(headers: Sequence[str], rows: Iterable[Sequence[object]]) -> None:
    values = [[str(value) for value in row] for row in rows]
    widths = [len(item) for item in headers]
    for row in values:
        widths = [max(width, len(row[index])) for index, width in enumerate(widths)]
    template = "  ".join(f"{{:<{width}}}" for width in widths)
    print(template.format(*headers))
    print(template.format(*(('-' * width) for width in widths)))
    for row in values:
        print(template.format(*row))


def print_services(services: list[ServiceInfo]) -> None:
    print_table(
        ("SERVICE", "STATE", "START TYPE", "PID", "DISPLAY NAME"),
        (
            (item.name, item.state, item.start_type, item.pid or "-", item.display_name)
            for item in services
        ),
    )
    print(f"\nTotal services: {len(services)}")


def print_changes(changes: list[ServiceChange]) -> None:
    if not changes:
        print("No service state, startup type, or PID changes were detected.")
        return
    print_table(
        ("SERVICE", "KIND", "STATE", "START TYPE", "PID"),
        (
            (
                item.name,
                item.kind,
                f"{item.old_state or '-'} -> {item.new_state or '-'}",
                f"{item.old_start_type or '-'} -> {item.new_start_type or '-'}",
                f"{item.old_pid if item.old_pid is not None else '-'} -> "
                f"{item.new_pid if item.new_pid is not None else '-'}",
            )
            for item in changes
        ),
    )


def command_live(_: argparse.Namespace) -> int:
    print_services(enumerate_services())
    return 0


def command_snapshot(args: argparse.Namespace) -> int:
    services = enumerate_services()
    with SnapshotStore(args.db) as store:
        snapshot = store.save_snapshot(services, label=args.label)
    print(
        f"Saved snapshot {snapshot.id} with {len(services)} services at "
        f"{snapshot.captured_at}."
    )
    return 0


def command_snapshots(args: argparse.Namespace) -> int:
    with SnapshotStore(args.db) as store:
        snapshots = store.list_snapshots()
    print_table(
        ("ID", "CAPTURED AT", "LABEL"),
        ((item.id, item.captured_at, item.label or "") for item in snapshots),
    )
    return 0


def command_diff(args: argparse.Namespace) -> int:
    with SnapshotStore(args.db) as store:
        if args.before is None or args.after is None:
            ids = store.latest_ids(2)
            if len(ids) < 2:
                raise ValueError("At least two snapshots are required")
            before_id, after_id = ids
        else:
            before_id, after_id = args.before, args.after
        changes = diff_services(
            store.get_services(before_id), store.get_services(after_id)
        )
    print(f"Snapshot {before_id} -> snapshot {after_id}\n")
    print_changes(changes)
    return 0


def command_groups(args: argparse.Namespace) -> int:
    if args.snapshot is None:
        services = enumerate_services()
        source = "live data"
    else:
        with SnapshotStore(args.db) as store:
            services = store.get_services(args.snapshot)
        source = f"snapshot {args.snapshot}"
    groups = group_by_pid(services, args.minimum)
    print(f"Shared service hosts from {source}\n")
    for pid, items in groups[: args.limit]:
        print(f"PID {pid}: {len(items)} services")
        print("  " + ", ".join(item.name for item in items))
    if not groups:
        print("No PIDs matched the requested minimum service count.")
    return 0


def command_watch(args: argparse.Namespace) -> int:
    previous = enumerate_services()
    print(f"Watching {len(previous)} services. Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(args.interval)
            current = enumerate_services()
            changes = diff_services(previous, current)
            if changes:
                print(f"\nChanges detected at {time.strftime('%Y-%m-%d %H:%M:%S')}")
                print_changes(changes)
            previous = current
    except KeyboardInterrupt:
        print("\nWatch stopped.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Query Windows services through the raw Service Control Manager API."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    live = subparsers.add_parser("live", help="print the current service list")
    live.set_defaults(handler=command_live)

    snapshot = subparsers.add_parser("snapshot", help="save a service snapshot")
    snapshot.add_argument("--db", default="services.db", help="SQLite database path")
    snapshot.add_argument("--label", help="optional snapshot label")
    snapshot.set_defaults(handler=command_snapshot)

    snapshots = subparsers.add_parser("snapshots", help="list saved snapshots")
    snapshots.add_argument("--db", default="services.db", help="SQLite database path")
    snapshots.set_defaults(handler=command_snapshots)

    diff = subparsers.add_parser("diff", help="compare two saved snapshots")
    diff.add_argument("before", nargs="?", type=int, help="earlier snapshot ID")
    diff.add_argument("after", nargs="?", type=int, help="later snapshot ID")
    diff.add_argument("--db", default="services.db", help="SQLite database path")
    diff.set_defaults(handler=command_diff)

    groups = subparsers.add_parser("groups", help="show PIDs hosting multiple services")
    groups.add_argument("--snapshot", type=int, help="use a saved snapshot instead of live data")
    groups.add_argument("--db", default="services.db", help="SQLite database path")
    groups.add_argument("--minimum", type=int, default=2, help="minimum services per PID")
    groups.add_argument("--limit", type=int, default=10, help="maximum groups to display")
    groups.set_defaults(handler=command_groups)

    watch = subparsers.add_parser("watch", help="print live changes continuously")
    watch.add_argument("--interval", type=float, default=3.0, help="poll interval in seconds")
    watch.set_defaults(handler=command_watch)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "diff" and ((args.before is None) != (args.after is None)):
        parser.error("provide both snapshot IDs or neither")
    try:
        return int(args.handler(args))
    except (OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
