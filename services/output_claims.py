from __future__ import annotations

import threading
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

_OUTPUT_LOCKS: dict[str, list[object]] = {}
_OUTPUT_LOCKS_GUARD = threading.Lock()


@contextmanager
def output_lock(path: Path) -> Iterator[None]:
    """Serialize same-path jobs in this process, including explicit overwrite."""
    key = str(path).lower() if path.drive else str(path)
    with _OUTPUT_LOCKS_GUARD:
        entry = _OUTPUT_LOCKS.setdefault(key, [threading.Lock(), 0])
        lock, _ = entry
        entry[1] += 1
    with lock:
        try:
            yield
        finally:
            with _OUTPUT_LOCKS_GUARD:
                entry[1] -= 1
                if entry[1] == 0:
                    _OUTPUT_LOCKS.pop(key, None)


@contextmanager
def claim_output(path: Path, if_exists: str) -> Iterator[bool]:
    """Reserve a new destination, yielding whether an existing output was skipped."""
    with output_lock(path):
        skipped = False
        if if_exists == "skip" and path.exists():
            skipped = True
        else:
            if path.is_dir():
                raise IsADirectoryError(str(path))
            if if_exists != "overwrite":
                try:
                    with path.open("xb"):
                        pass
                except FileExistsError:
                    if if_exists == "skip":
                        skipped = True
                    else:
                        raise
        try:
            yield skipped
        finally:
            if not skipped and if_exists != "overwrite":
                release_output(path, if_exists)


def release_output(path: Path, if_exists: str) -> None:
    """Remove only the empty reservation, never a committed document."""
    try:
        with path.open("rb") as stream:
            stream.seek(0, 2)
            if stream.tell() != 0:
                return
        path.unlink()
    except FileNotFoundError:
        pass
