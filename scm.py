"""Low-level ctypes wrapper around the Windows Service Control Manager API."""

from __future__ import annotations

import ctypes
import os
from ctypes import wintypes

from models import ServiceInfo


if os.name != "nt":
    raise RuntimeError("scm.py can only run on Windows")


advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)

SC_HANDLE = wintypes.HANDLE
SC_MANAGER_CONNECT = 0x0001
SC_MANAGER_ENUMERATE_SERVICE = 0x0004
SERVICE_QUERY_CONFIG = 0x0001
SERVICE_WIN32 = 0x00000030
SERVICE_STATE_ALL = 0x00000003
SC_ENUM_PROCESS_INFO = 0
SERVICE_CONFIG_DELAYED_AUTO_START_INFO = 3
ERROR_ACCESS_DENIED = 5
ERROR_INSUFFICIENT_BUFFER = 122
ERROR_MORE_DATA = 234

STATE_NAMES = {
    1: "STOPPED",
    2: "START_PENDING",
    3: "STOP_PENDING",
    4: "RUNNING",
    5: "CONTINUE_PENDING",
    6: "PAUSE_PENDING",
    7: "PAUSED",
}

START_TYPE_NAMES = {
    0: "BOOT",
    1: "SYSTEM",
    2: "AUTOMATIC",
    3: "MANUAL",
    4: "DISABLED",
}


class SERVICE_STATUS_PROCESS(ctypes.Structure):
    _fields_ = [
        ("dwServiceType", wintypes.DWORD),
        ("dwCurrentState", wintypes.DWORD),
        ("dwControlsAccepted", wintypes.DWORD),
        ("dwWin32ExitCode", wintypes.DWORD),
        ("dwServiceSpecificExitCode", wintypes.DWORD),
        ("dwCheckPoint", wintypes.DWORD),
        ("dwWaitHint", wintypes.DWORD),
        ("dwProcessId", wintypes.DWORD),
        ("dwServiceFlags", wintypes.DWORD),
    ]


class ENUM_SERVICE_STATUS_PROCESSW(ctypes.Structure):
    _fields_ = [
        ("lpServiceName", wintypes.LPWSTR),
        ("lpDisplayName", wintypes.LPWSTR),
        ("ServiceStatusProcess", SERVICE_STATUS_PROCESS),
    ]


class QUERY_SERVICE_CONFIGW(ctypes.Structure):
    _fields_ = [
        ("dwServiceType", wintypes.DWORD),
        ("dwStartType", wintypes.DWORD),
        ("dwErrorControl", wintypes.DWORD),
        ("lpBinaryPathName", wintypes.LPWSTR),
        ("lpLoadOrderGroup", wintypes.LPWSTR),
        ("dwTagId", wintypes.DWORD),
        ("lpDependencies", wintypes.LPWSTR),
        ("lpServiceStartName", wintypes.LPWSTR),
        ("lpDisplayName", wintypes.LPWSTR),
    ]


class SERVICE_DELAYED_AUTO_START_INFO(ctypes.Structure):
    _fields_ = [("fDelayedAutostart", wintypes.BOOL)]


advapi32.OpenSCManagerW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD]
advapi32.OpenSCManagerW.restype = SC_HANDLE
advapi32.EnumServicesStatusExW.argtypes = [
    SC_HANDLE,
    ctypes.c_int,
    wintypes.DWORD,
    wintypes.DWORD,
    ctypes.POINTER(ctypes.c_ubyte),
    wintypes.DWORD,
    ctypes.POINTER(wintypes.DWORD),
    ctypes.POINTER(wintypes.DWORD),
    ctypes.POINTER(wintypes.DWORD),
    wintypes.LPCWSTR,
]
advapi32.EnumServicesStatusExW.restype = wintypes.BOOL
advapi32.OpenServiceW.argtypes = [SC_HANDLE, wintypes.LPCWSTR, wintypes.DWORD]
advapi32.OpenServiceW.restype = SC_HANDLE
advapi32.QueryServiceConfigW.argtypes = [
    SC_HANDLE,
    ctypes.POINTER(QUERY_SERVICE_CONFIGW),
    wintypes.DWORD,
    ctypes.POINTER(wintypes.DWORD),
]
advapi32.QueryServiceConfigW.restype = wintypes.BOOL
advapi32.QueryServiceConfig2W.argtypes = [
    SC_HANDLE,
    wintypes.DWORD,
    ctypes.POINTER(ctypes.c_ubyte),
    wintypes.DWORD,
    ctypes.POINTER(wintypes.DWORD),
]
advapi32.QueryServiceConfig2W.restype = wintypes.BOOL
advapi32.CloseServiceHandle.argtypes = [SC_HANDLE]
advapi32.CloseServiceHandle.restype = wintypes.BOOL


def _raise_last_error(action: str) -> None:
    code = ctypes.get_last_error()
    raise OSError(code, f"{action}: {ctypes.FormatError(code).strip()}")


def _open_scm() -> SC_HANDLE:
    handle = advapi32.OpenSCManagerW(
        None, None, SC_MANAGER_CONNECT | SC_MANAGER_ENUMERATE_SERVICE
    )
    if not handle:
        _raise_last_error("OpenSCManagerW failed")
    return handle


def _enumerate_statuses(scm_handle: SC_HANDLE) -> list[tuple[str, str, str, int]]:
    needed = wintypes.DWORD(0)
    count = wintypes.DWORD(0)
    resume = wintypes.DWORD(0)

    advapi32.EnumServicesStatusExW(
        scm_handle,
        SC_ENUM_PROCESS_INFO,
        SERVICE_WIN32,
        SERVICE_STATE_ALL,
        None,
        0,
        ctypes.byref(needed),
        ctypes.byref(count),
        ctypes.byref(resume),
        None,
    )
    error = ctypes.get_last_error()
    if error not in (ERROR_MORE_DATA, ERROR_INSUFFICIENT_BUFFER):
        _raise_last_error("EnumServicesStatusExW sizing call failed")

    # The sizing probe must not advance the real enumeration cursor.
    resume.value = 0

    while True:
        buffer = (ctypes.c_ubyte * needed.value)()
        ok = advapi32.EnumServicesStatusExW(
            scm_handle,
            SC_ENUM_PROCESS_INFO,
            SERVICE_WIN32,
            SERVICE_STATE_ALL,
            buffer,
            needed.value,
            ctypes.byref(needed),
            ctypes.byref(count),
            ctypes.byref(resume),
            None,
        )
        if ok:
            break
        if ctypes.get_last_error() != ERROR_MORE_DATA:
            _raise_last_error("EnumServicesStatusExW failed")

    array_type = ENUM_SERVICE_STATUS_PROCESSW * count.value
    entries = ctypes.cast(buffer, ctypes.POINTER(array_type)).contents
    result: list[tuple[str, str, str, int]] = []
    for entry in entries:
        status = entry.ServiceStatusProcess
        result.append(
            (
                entry.lpServiceName,
                entry.lpDisplayName,
                STATE_NAMES.get(status.dwCurrentState, f"UNKNOWN_{status.dwCurrentState}"),
                int(status.dwProcessId),
            )
        )
    return result


def _query_start_type(scm_handle: SC_HANDLE, service_name: str) -> str:
    service = advapi32.OpenServiceW(scm_handle, service_name, SERVICE_QUERY_CONFIG)
    if not service:
        error = ctypes.get_last_error()
        if error == ERROR_ACCESS_DENIED:
            return "ACCESS_DENIED"
        return f"ERROR_{error}"

    try:
        needed = wintypes.DWORD(0)
        advapi32.QueryServiceConfigW(service, None, 0, ctypes.byref(needed))
        error = ctypes.get_last_error()
        if error != ERROR_INSUFFICIENT_BUFFER:
            return f"ERROR_{error}"

        buffer = (ctypes.c_ubyte * needed.value)()
        config = ctypes.cast(buffer, ctypes.POINTER(QUERY_SERVICE_CONFIGW))
        if not advapi32.QueryServiceConfigW(
            service, config, needed.value, ctypes.byref(needed)
        ):
            return f"ERROR_{ctypes.get_last_error()}"

        start_type = START_TYPE_NAMES.get(
            config.contents.dwStartType, f"UNKNOWN_{config.contents.dwStartType}"
        )
        if start_type == "AUTOMATIC" and _is_delayed_auto_start(service):
            return "AUTOMATIC_DELAYED"
        return start_type
    finally:
        advapi32.CloseServiceHandle(service)


def _is_delayed_auto_start(service_handle: SC_HANDLE) -> bool:
    info = SERVICE_DELAYED_AUTO_START_INFO()
    needed = wintypes.DWORD(0)
    buffer = ctypes.cast(ctypes.byref(info), ctypes.POINTER(ctypes.c_ubyte))
    ok = advapi32.QueryServiceConfig2W(
        service_handle,
        SERVICE_CONFIG_DELAYED_AUTO_START_INFO,
        buffer,
        ctypes.sizeof(info),
        ctypes.byref(needed),
    )
    return bool(ok and info.fDelayedAutostart)


def enumerate_services() -> list[ServiceInfo]:
    """Return every registered Win32 service with status, startup type, and PID."""

    scm_handle = _open_scm()
    try:
        statuses = _enumerate_statuses(scm_handle)
        services = [
            ServiceInfo(
                name=name,
                display_name=display_name,
                state=state,
                start_type=_query_start_type(scm_handle, name),
                pid=pid,
            )
            for name, display_name, state, pid in statuses
        ]
        return sorted(services, key=lambda item: item.name.casefold())
    finally:
        advapi32.CloseServiceHandle(scm_handle)
