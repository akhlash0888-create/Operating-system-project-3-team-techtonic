# Windows Services Monitor

This project queries the Windows Service Control Manager through the raw Win32 API by using Python `ctypes`. It lists each registered service, its live state, configured startup type, and hosting process ID. It can store snapshots in SQLite, compare snapshots, group services by PID, and watch for changes.

## Requirements

- Windows 10 or Windows 11
- Python 3.10 or newer
- Standard user access is enough for normal queries

The program uses only the Python standard library at runtime.

## Files

- `scm.py` - calls `OpenSCManagerW`, `EnumServicesStatusExW`, `OpenServiceW`, `QueryServiceConfigW`, and `QueryServiceConfig2W`
- `models.py` - typed data models
- `db.py` - SQLite snapshot storage
- `analysis.py` - snapshot diff and PID grouping
- `main.py` - command-line interface
- `tests/test_project.py` - automated tests

## Commands

```powershell
python main.py live
python main.py snapshot --label before
python main.py snapshot --label after
python main.py snapshots
python main.py diff
python main.py groups --limit 10
python main.py watch --interval 3
```

Use a different database with `--db path\to\file.db`. The `diff` command compares the latest two snapshots when IDs are omitted, or accepts two explicit IDs:

```powershell
python main.py diff 1 2
```

## Safe state-change demo

1. Save a `before` snapshot.
2. In `services.msc`, change only a service that you own or have permission to operate on.
3. Wait until the state settles.
4. Save an `after` snapshot.
5. Run `python main.py diff`.
6. Restore the service to its original state if the demonstration changed it.

The monitor itself never starts, stops, or reconfigures a service.

## Tests

```powershell
python -m unittest discover -s tests -v
```

The tests cover state/start-type/PID changes, added and removed services, shared-PID grouping, SQLite round trips, and missing snapshot handling.
