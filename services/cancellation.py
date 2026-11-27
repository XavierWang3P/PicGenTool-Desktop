from __future__ import annotations

import threading
from collections.abc import Callable
from typing import TypeVar


T = TypeVar("T")


class GenerationCancelled(Exception):
    def __init__(self, message: str = "生成已取消") -> None:
        super().__init__(message)


class CancellationToken:
    def __init__(self, event: threading.Event | None = None) -> None:
        self._event = event or threading.Event()
        self._lock = threading.Lock()
        self._committed = False

    def cancel(self) -> None:
        with self._lock:
            if not self._committed:
                self._event.set()

    def is_cancelled(self) -> bool:
        return self._event.is_set()

    def raise_if_cancelled(self) -> None:
        with self._lock:
            if self._event.is_set():
                raise GenerationCancelled()

    def commit(self, action: Callable[[], T]) -> T:
        """Run the final atomic action unless cancellation already won the race."""
        with self._lock:
            if self._event.is_set():
                raise GenerationCancelled()
            result = action()
            self._committed = True
            return result
