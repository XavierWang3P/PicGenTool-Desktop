from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from models.agent_contract import ERROR_EXIT_CODES, request_schema
from services.agent_api import AgentAPI

MAX_INPUT_BYTES = 1024 * 1024
AGENT_OPTIONS = {
    "--generate",
    "--validate",
    "--capabilities",
    "--schema",
    "--serve",
    "--batch",
}


class _ArgumentParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        raise ValueError(message)


def build_parser() -> argparse.ArgumentParser:
    parser = _ArgumentParser(description="PicGenTool 无界面自动化接口")
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--generate", metavar="JSON_FILE", help="生成文档，- 表示 stdin")
    action.add_argument("--validate", metavar="JSON_FILE", help="只校验请求，不生成")
    action.add_argument("--capabilities", action="store_true", help="报告支持能力")
    action.add_argument("--schema", action="store_true", help="输出请求 JSON Schema")
    action.add_argument("--diagnose", action="store_true", help="输出机器可读诊断")
    action.add_argument("--serve", action="store_true", help="运行 stdin/stdout NDJSON 服务")
    action.add_argument("--batch", metavar="JSONL_FILE", help="批量处理 JSON Lines 请求")
    parser.add_argument("--json", action="store_true", help="使用 JSON 输出（自动化默认）")
    parser.add_argument("--workers", type=int, default=1, choices=range(1, 5))
    parser.add_argument("--on-error", choices=("continue", "stop"), default="continue")
    parser.add_argument("--timeout", type=float, default=None)
    parser.add_argument("--result-file", type=Path)
    return parser


def _configure_utf8(stream) -> None:
    reconfigure = getattr(stream, "reconfigure", None)
    if reconfigure is not None:
        reconfigure(encoding="utf-8")


def _diagnostics_payload() -> dict:
    from services.diagnostics import DiagnosticsService

    snapshot = DiagnosticsService().collect(deep_templates=True)
    template = snapshot.template_result
    return {
        "ok": template.valid and snapshot.heif_available,
        "app_version": snapshot.app_version,
        "python_version": snapshot.python_version,
        "platform": snapshot.platform,
        "packaged": snapshot.packaged,
        "heif": {
            "available": snapshot.heif_available,
            "version": snapshot.heif_version,
        },
        "templates": {
            "valid": template.valid,
            "checked_count": template.checked_count,
            "issues": [
                {
                    "image_count": issue.template_number,
                    "code": issue.code,
                    "message": issue.message,
                }
                for issue in template.issues
            ],
        },
        "log_available": snapshot.log_path is not None,
        "explorer_menu": snapshot.explorer_status.value,
    }


def _read_request(source: str, stdin) -> object:
    try:
        if source == "-":
            raw = stdin.buffer.read(MAX_INPUT_BYTES + 1)
        else:
            path = Path(source)
            with path.open("rb") as request_file:
                raw = request_file.read(MAX_INPUT_BYTES + 1)
        if len(raw) > MAX_INPUT_BYTES:
            raise ValueError("请求超过 1 MiB 限制")
        return json.loads(raw.decode("utf-8-sig"))
    except ValueError:
        raise
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("无法读取有效 UTF-8 JSON 请求") from exc


def _write_json(payload: object, stdout) -> None:
    stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    stdout.flush()


def _exit_code(result: dict) -> int:
    if result.get("ok"):
        return 0
    return ERROR_EXIT_CODES.get(result.get("error", {}).get("code"), 1)


def run_agent_cli(argv: list[str], stdin=None, stdout=None, stderr=None) -> int:
    stdin = stdin or sys.stdin
    stdout = stdout or sys.stdout
    stderr = stderr or sys.stderr
    _configure_utf8(stdin)
    _configure_utf8(stdout)
    _configure_utf8(stderr)
    try:
        args = build_parser().parse_args(argv)
        api = AgentAPI()
        if args.diagnose:
            result = _diagnostics_payload()
            code = 0 if result["ok"] else 3
        elif args.schema:
            result = request_schema()
            code = 0
        elif args.capabilities:
            result = api.capabilities()
            code = 0 if result["ok"] else 3
        elif args.validate or args.generate:
            payload = _read_request(args.validate or args.generate, stdin)
            result = api.validate(payload) if args.validate else api.generate(payload)
            code = _exit_code(result)
        elif args.batch or args.serve:
            if args.result_file:
                raise ValueError("批处理和常驻服务不支持 --result-file")
            from services.agent_runner import run_batch_file, run_server

            if args.batch:
                results = run_batch_file(
                    Path(args.batch),
                    api,
                    workers=args.workers,
                    on_error=args.on_error,
                    timeout=args.timeout,
                )
                code = 0
                for result in results:
                    _write_json(result, stdout)
                    code = max(code, _exit_code(result))
                return code
            return run_server(api, stdin=stdin, stdout=stdout)
        else:
            raise ValueError("缺少操作")
    except ValueError as exc:
        result = api.failure(api_error(str(exc))) if "api" in locals() else {
            "ok": False,
            "error": {"code": "INVALID_REQUEST", "message": str(exc), "details": {}},
        }
        code = 2
    except Exception as exc:
        result = api.failure(exc) if "api" in locals() else {
            "ok": False,
            "error": {"code": "INTERNAL_ERROR", "message": "发生未预期错误", "details": {}},
        }
        code = 1
    if args.result_file if "args" in locals() else False:
        args.result_file.write_text(
            json.dumps(result, ensure_ascii=False), encoding="utf-8"
        )
    _write_json(result, stdout)
    return code


def api_error(message: str):
    from services.agent_api import AgentError

    return AgentError("INVALID_REQUEST", message)


if __name__ == "__main__":
    raise SystemExit(run_agent_cli(sys.argv[1:]))
