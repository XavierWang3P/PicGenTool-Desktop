from __future__ import annotations

import json
import threading
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Iterable, Iterator

from services.agent_api import AgentAPI
from services.cancellation import CancellationToken

MAX_LINE_BYTES = 1024 * 1024


def read_jsonl(path: Path) -> Iterator[tuple[int, object]]:
    """Yield one-based input line numbers and payloads without hiding bad lines."""
    with path.open("rb") as source:
        for number, raw in enumerate(source, start=1):
            if len(raw) > MAX_LINE_BYTES:
                yield number, AgentLineError("请求行超过 1 MiB 限制")
                continue
            try:
                text = raw.decode("utf-8-sig").strip()
            except UnicodeDecodeError:
                yield number, AgentLineError("该行不是有效 UTF-8 JSON")
                continue
            if not text or text.startswith("#"):
                continue
            try:
                yield number, json.loads(text)
            except (UnicodeDecodeError, json.JSONDecodeError):
                yield number, AgentLineError("该行不是有效 UTF-8 JSON")


class AgentLineError(Exception):
    pass


def _run_one(
    api: AgentAPI,
    payload: object,
    timeout: float | None,
    tokens: set[CancellationToken] | None = None,
    tokens_lock: threading.Lock | None = None,
) -> dict:
    if isinstance(payload, AgentLineError):
        return api.failure(api_error(str(payload)))
    token = CancellationToken()
    if tokens is not None and tokens_lock is not None:
        with tokens_lock:
            tokens.add(token)
    timer = None
    if timeout is not None:
        if timeout <= 0:
            return api.failure(api_error("timeout 必须大于 0 秒"))
        timer = threading.Timer(timeout, token.cancel)
        timer.daemon = True
        timer.start()
    try:
        return api.generate(payload, cancellation=token)
    finally:
        if timer is not None:
            timer.cancel()
        if tokens is not None and tokens_lock is not None:
            with tokens_lock:
                tokens.discard(token)


def run_batch(
    payloads: Iterable[tuple[int, object]],
    api: AgentAPI,
    workers: int = 1,
    on_error: str = "continue",
    timeout: float | None = None,
) -> Iterator[dict]:
    """Submit a bounded queue; preserve input order in the result stream."""
    if workers not in range(1, 5):
        raise ValueError("workers 必须为 1 到 4")
    if on_error not in {"continue", "stop"}:
        raise ValueError("on_error 必须为 continue 或 stop")
    source = iter(payloads)
    stopped = False
    tokens: set[CancellationToken] = set()
    tokens_lock = threading.Lock()
    with ThreadPoolExecutor(max_workers=workers, thread_name_prefix="pgt-agent") as pool:
        pending: list[tuple[int, Future]] = []
        while True:
            while not stopped and len(pending) < workers:
                try:
                    item = next(source)
                    pending.append(
                        (
                            item[0],
                            pool.submit(
                                _run_one, api, item[1], timeout, tokens, tokens_lock
                            ),
                        )
                    )
                except StopIteration:
                    break
            if not pending:
                break
            line, future = pending.pop(0)
            try:
                result = future.result()
            except KeyboardInterrupt:
                with tokens_lock:
                    for token in tokens:
                        token.cancel()
                for _, running in pending:
                    running.cancel()
                raise
            yield {"line": line, **result}
            if on_error == "stop" and not result.get("ok"):
                stopped = True


def run_batch_file(path: Path, api: AgentAPI, **options) -> Iterator[dict]:
    try:
        yield from run_batch(read_jsonl(path), api, **options)
    except OSError:
        yield api.failure(api_error("无法读取批处理文件"))


def run_server(api: AgentAPI, stdin, stdout) -> int:
    """Serve correlated one-job requests until EOF, one JSON object per line."""
    for raw in stdin.buffer:
        request_id = None
        try:
            if len(raw) > MAX_LINE_BYTES:
                raise ValueError("请求行超过 1 MiB 限制")
            envelope = json.loads(raw.decode("utf-8-sig"))
            if not isinstance(envelope, dict) or set(envelope) - {"id", "request"}:
                raise ValueError("服务请求必须只包含 id 和 request")
            request_id = envelope.get("id")
            if "request" not in envelope:
                raise ValueError("缺少 request")
            result = api.generate(envelope["request"])
        except KeyboardInterrupt:
            raise
        except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
            result = api.failure(api_error(str(exc)))
        stdout.write(json.dumps({"id": request_id, "result": result}, ensure_ascii=False) + "\n")
        stdout.flush()
    return 0


def api_error(message: str):
    from services.agent_api import AgentError

    return AgentError("INVALID_REQUEST", message)
