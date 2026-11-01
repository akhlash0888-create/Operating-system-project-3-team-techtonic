"""Typed data models used by the Services Monitor project."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ServiceInfo:
    """Runtime and configured information for one Windows service."""

    name: str
    display_name: str
    state: str
    start_type: str
    pid: int


@dataclass(frozen=True, slots=True)
class Snapshot:
    """Metadata describing a stored service snapshot."""

    id: int
    captured_at: str
    label: str | None


@dataclass(frozen=True, slots=True)
class ServiceChange:
    """The difference for a service between two snapshots."""

    name: str
    display_name: str
    old_state: str | None
    new_state: str | None
    old_start_type: str | None
    new_start_type: str | None
    old_pid: int | None
    new_pid: int | None

    @property
    def kind(self) -> str:
        if self.old_state is None:
            return "ADDED"
        if self.new_state is None:
            return "REMOVED"
        return "CHANGED"
