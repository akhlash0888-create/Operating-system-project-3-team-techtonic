"""Snapshot comparison and shared-process analysis."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable

from models import ServiceChange, ServiceInfo


def diff_services(
    before: Iterable[ServiceInfo], after: Iterable[ServiceInfo]
) -> list[ServiceChange]:
    old = {item.name.casefold(): item for item in before}
    new = {item.name.casefold(): item for item in after}
    changes: list[ServiceChange] = []

    for key in sorted(old.keys() | new.keys()):
        left = old.get(key)
        right = new.get(key)
        if left == right:
            continue
        if left and right and (
            left.state,
            left.start_type,
            left.pid,
        ) == (
            right.state,
            right.start_type,
            right.pid,
        ):
            continue
        sample = right or left
        assert sample is not None
        changes.append(
            ServiceChange(
                name=sample.name,
                display_name=sample.display_name,
                old_state=left.state if left else None,
                new_state=right.state if right else None,
                old_start_type=left.start_type if left else None,
                new_start_type=right.start_type if right else None,
                old_pid=left.pid if left else None,
                new_pid=right.pid if right else None,
            )
        )
    return changes


def group_by_pid(
    services: Iterable[ServiceInfo], minimum_services: int = 2
) -> list[tuple[int, list[ServiceInfo]]]:
    groups: dict[int, list[ServiceInfo]] = defaultdict(list)
    for service in services:
        if service.pid > 0:
            groups[service.pid].append(service)

    filtered = [
        (pid, sorted(items, key=lambda item: item.name.casefold()))
        for pid, items in groups.items()
        if len(items) >= minimum_services
    ]
    return sorted(filtered, key=lambda pair: (-len(pair[1]), pair[0]))
