from __future__ import annotations

import logging
import sys
import threading
import traceback as traceback_module
from collections.abc import Iterable
from pathlib import Path
from types import TracebackType
from typing import Callable

from PyQt6.QtCore import QObject, pyqtSignal


SysHook = Callable[[type[BaseException], BaseException, TracebackType | None], None]
ThreadHook = Callable[[threading.ExceptHookArgs], None]


class GlobalExceptionHandler(QObject):
    notification = pyqtSignal(str)

    def __init__(
        self,
        logger: logging.Logger,
        *,
        original_sys_hook: SysHook | None = None,
        original_thread_hook: ThreadHook | None = None,
    ) -> None:
        super().__init__()
        self.logger = logger
        self.original_sys_hook = original_sys_hook or sys.__excepthook__
        self.original_thread_hook = original_thread_hook or threading.__excepthook__
        self._sensitive_values: tuple[str, ...] = ()
        self._handling = threading.local()

    def set_sensitive_values(self, values: Iterable[str | Path]) -> None:
        candidates: set[str] = set()
        for value in values:
            text = str(value).strip()
            if text:
                candidates.add(text)
            if isinstance(value, Path):
                try:
                    candidates.add(str(value.resolve()))
                except OSError:
                    pass
        self._sensitive_values = tuple(sorted(candidates, key=len, reverse=True))

    def clear_sensitive_values(self) -> None:
        self._sensitive_values = ()

    def capture(
        self,
        exc_type: type[BaseException],
        exc_value: BaseException,
        traceback: TracebackType | None,
    ) -> None:
        if issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
            self._call_sys_fallback(exc_type, exc_value, traceback)
            return
        if getattr(self._handling, "active", False):
            self._call_sys_fallback(exc_type, exc_value, traceback)
            return

        self._handling.active = True
        try:
            formatted = "".join(
                traceback_module.format_exception(exc_type, exc_value, traceback)
            )
            self.logger.critical("未捕获异常：\n%s", self._sanitize(formatted))
            self.notification.emit("发生未处理异常。请在“关于 / 诊断”中复制诊断信息。")
        except Exception:
            self._call_sys_fallback(exc_type, exc_value, traceback)
        finally:
            self._handling.active = False

    def capture_thread(self, args: threading.ExceptHookArgs) -> None:
        if issubclass(args.exc_type, (KeyboardInterrupt, SystemExit)):
            self._call_thread_fallback(args)
            return
        if getattr(self._handling, "active", False):
            self._call_thread_fallback(args)
            return

        self._handling.active = True
        try:
            formatted = "".join(
                traceback_module.format_exception(
                    args.exc_type,
                    args.exc_value,
                    args.exc_traceback,
                )
            )
            thread_name = getattr(args.thread, "name", "unknown")
            self.logger.critical(
                "后台线程未捕获异常（%s）：\n%s",
                thread_name,
                self._sanitize(formatted),
            )
            self.notification.emit("后台任务发生未处理异常。请在“关于 / 诊断”中复制诊断信息。")
        except Exception:
            self._call_thread_fallback(args)
        finally:
            self._handling.active = False

    def _sanitize(self, text: str) -> str:
        home = str(Path.home())
        replacements = [*self._sensitive_values, home]
        sanitized = text
        for value in sorted(set(replacements), key=len, reverse=True):
            if not value:
                continue
            replacement = "<user-home>" if value == home else "<redacted>"
            sanitized = sanitized.replace(value, replacement)
        return sanitized

    def _call_sys_fallback(
        self,
        exc_type: type[BaseException],
        exc_value: BaseException,
        traceback: TracebackType | None,
    ) -> None:
        try:
            self.original_sys_hook(exc_type, exc_value, traceback)
        except Exception:
            pass

    def _call_thread_fallback(self, args: threading.ExceptHookArgs) -> None:
        try:
            self.original_thread_hook(args)
        except Exception:
            pass


def install_exception_handlers(app, logger: logging.Logger) -> GlobalExceptionHandler:
    from PyQt6.QtCore import QCoreApplication
    from PyQt6.QtWidgets import QMessageBox

    handler = GlobalExceptionHandler(
        logger,
        original_sys_hook=sys.excepthook,
        original_thread_hook=threading.excepthook,
    )

    def show_dialog(message: str) -> None:
        if not QCoreApplication.closingDown():
            QMessageBox.critical(None, "应用异常", message)

    handler.notification.connect(show_dialog)
    sys.excepthook = handler.capture
    threading.excepthook = handler.capture_thread
    app._pgt_exception_handler = handler
    return handler
