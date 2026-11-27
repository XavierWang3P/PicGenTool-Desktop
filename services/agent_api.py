from __future__ import annotations

import logging
import time
from datetime import date
from pathlib import Path
from uuid import uuid4

from config import APP_VERSION, MAX_IMAGES, SUPPORTED_FORMATS
from models.agent_contract import SCHEMA_VERSION
from models.generation_task import GenerationTask
from services.cancellation import GenerationCancelled
from services.generation_service import GenerationService
from services.image_service import ImageService
from services.template_service import TemplateService

_REQUIRED_FIELDS = {
    "schema_version",
    "title",
    "location",
    "activity_date",
    "images",
    "output",
}
_ALLOWED_FIELDS = _REQUIRED_FIELDS | {"options"}
_OPTION_FIELDS = {"if_exists", "on_image_error", "open", "use_settings"}


class AgentError(Exception):
    def __init__(self, code: str, message: str, details: dict | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}


class AgentAPI:
    """Validate automation requests and expose the shared generation pipeline."""

    def __init__(
        self,
        image_service: ImageService | None = None,
        generation_service: GenerationService | None = None,
        template_service: TemplateService | None = None,
    ) -> None:
        self.image_service = image_service or ImageService()
        self.generation_service = generation_service or GenerationService(
            image_service=self.image_service
        )
        self.template_service = template_service or TemplateService()
        self.logger = logging.getLogger("pgt.agent")

    def validate(self, payload: object) -> dict:
        started = time.perf_counter()
        job_id = uuid4().hex
        try:
            prepared = self._prepare(payload)
            return self._success(
                job_id,
                prepared["task"].save_path,
                len(prepared["images"]),
                prepared["duplicates"],
                [],
                started,
                validated=True,
            )
        except Exception as exc:
            return self.failure(exc, job_id=job_id, started=started)

    def generate(
        self,
        payload: object,
        progress=None,
        cancellation=None,
    ) -> dict:
        started = time.perf_counter()
        job_id = uuid4().hex
        try:
            prepared = self._prepare(payload)
            task = prepared["task"]
            if prepared["skip"]:
                return self._success(
                    job_id,
                    task.save_path,
                    len(prepared["images"]),
                    prepared["duplicates"],
                    ["输出已存在，按要求跳过生成"],
                    started,
                    skipped=True,
                )
            output = self.generation_service.run(task, progress, cancellation)
            return self._success(
                job_id,
                output,
                len(prepared["images"]),
                prepared["duplicates"],
                [],
                started,
            )
        except Exception as exc:
            return self.failure(exc, job_id=job_id, started=started)

    def capabilities(self) -> dict:
        result = self.template_service.validate_all(deep=False)
        return {
            "ok": result.valid,
            "schema_version": SCHEMA_VERSION,
            "app_version": APP_VERSION,
            "max_images": MAX_IMAGES,
            "formats": sorted(SUPPORTED_FORMATS),
            "templates": {
                "available": result.checked_count,
                "missing": [
                    issue.template_number
                    for issue in result.issues
                    if issue.code == "missing_file"
                ],
            },
            "features": [
                "generate",
                "validate",
                "capabilities",
                "schema",
                "json",
                "batch",
                "serve",
                "cancel",
                "timeout",
            ],
        }

    def failure(self, error: Exception, job_id=None, started=None) -> dict:
        if isinstance(error, AgentError):
            code, message, details = error.code, error.message, error.details
        elif isinstance(error, GenerationCancelled):
            code, message, details = "CANCELLED", "生成已取消", {}
        else:
            code, message, details = self._map_error(error)
            self.logger.error("自动化请求失败: error_type=%s", type(error).__name__)
        result = {
            "ok": False,
            "job_id": job_id,
            "error": {"code": code, "message": message, "details": details},
            "app_version": APP_VERSION,
        }
        if started is not None:
            result["elapsed_ms"] = round((time.perf_counter() - started) * 1000, 3)
        return result

    def _prepare(self, payload: object) -> dict:
        data = self._validate_request(payload)
        output = Path(data["output"]).expanduser()
        if output.suffix.lower() != ".docx":
            raise AgentError("INVALID_REQUEST", "输出文件必须使用 .docx 扩展名")
        try:
            output = output.resolve(strict=False)
        except OSError as exc:
            raise AgentError("OUTPUT_NOT_WRITABLE", "输出路径无效") from exc

        input_paths: list[Path] = []
        for value in data["images"]:
            path = Path(value).expanduser()
            if not path.is_file():
                raise AgentError(
                    "IMAGE_NOT_FOUND",
                    "图片不存在或不是普通文件",
                    {"path": str(path)},
                )
            if path.suffix.lower() not in SUPPORTED_FORMATS:
                raise AgentError(
                    "UNSUPPORTED_FORMAT",
                    "不支持的图片格式",
                    {"path": str(path), "supported": sorted(SUPPORTED_FORMATS)},
                )
            input_paths.append(path)
        try:
            deduplicated = self.image_service.deduplicate(input_paths)
        except OSError as exc:
            raise AgentError("IMAGE_READ_FAILED", "无法读取图片") from exc
        if deduplicated.failures:
            raise AgentError(
                "IMAGE_READ_FAILED",
                "部分图片无法读取",
                {
                    "files": [
                        {"path": failure.filename, "reason": failure.reason}
                        for failure in deduplicated.failures
                    ]
                },
            )
        images = deduplicated.paths
        if not images:
            raise AgentError("INVALID_REQUEST", "去重后没有可用图片")
        if len(images) > MAX_IMAGES:
            raise AgentError(
                "TOO_MANY_IMAGES",
                f"最多可选择 {MAX_IMAGES} 张图片",
                {"image_count": len(images)},
            )
        try:
            self.template_service.get_template_path(len(images))
        except (ValueError, FileNotFoundError) as exc:
            raise AgentError(
                "TEMPLATE_MISSING",
                str(exc),
                {"image_count": len(images)},
            ) from exc

        if_exists = data["options"]["if_exists"]
        if output.exists():
            if if_exists == "skip":
                skip = True
            elif if_exists == "fail":
                raise AgentError(
                    "OUTPUT_EXISTS",
                    "输出文件已存在；可设置 options.if_exists=overwrite 或 skip",
                    {"output": str(output)},
                )
            else:
                skip = False
        else:
            skip = False
        if not skip:
            self._validate_output(output)
        task = GenerationTask(
            title=data["title"].strip(),
            location=data["location"].strip(),
            activity_date=date.fromisoformat(data["activity_date"]),
            image_paths=images,
            save_path=output,
            open_file=False,
        )
        return {
            "task": task,
            "images": images,
            "duplicates": deduplicated.duplicate_count,
            "skip": skip,
        }

    def _validate_request(self, payload: object) -> dict:
        if not isinstance(payload, dict):
            raise AgentError("INVALID_REQUEST", "请求必须是 JSON 对象")
        unknown = sorted(set(payload) - _ALLOWED_FIELDS)
        missing = sorted(_REQUIRED_FIELDS - set(payload))
        if unknown or missing:
            raise AgentError(
                "INVALID_REQUEST",
                "请求包含未知字段或缺少必要字段",
                {"unknown_fields": unknown, "missing_fields": missing},
            )
        if payload["schema_version"] != SCHEMA_VERSION or isinstance(
            payload["schema_version"], bool
        ):
            raise AgentError(
                "INVALID_REQUEST",
                f"schema_version 必须为 {SCHEMA_VERSION}",
            )
        for key in ("title", "location"):
            if not isinstance(payload[key], str) or not payload[key].strip():
                raise AgentError("INVALID_REQUEST", f"{key} 必须是非空字符串")
        if not isinstance(payload["output"], str) or not payload["output"].strip():
            raise AgentError("INVALID_REQUEST", "output 必须是非空字符串")
        if not isinstance(payload["activity_date"], str):
            raise AgentError("INVALID_REQUEST", "activity_date 必须是 YYYY-MM-DD 字符串")
        try:
            date.fromisoformat(payload["activity_date"])
        except ValueError as exc:
            raise AgentError(
                "INVALID_REQUEST", "activity_date 必须是有效的 YYYY-MM-DD 日期"
            ) from exc
        if (
            not isinstance(payload["images"], list)
            or not payload["images"]
            or not all(
                isinstance(value, str) and value.strip()
                for value in payload["images"]
            )
        ):
            raise AgentError("INVALID_REQUEST", "images 必须是非空图片路径字符串数组")
        if len(payload["images"]) > MAX_IMAGES:
            raise AgentError(
                "TOO_MANY_IMAGES",
                f"最多可提交 {MAX_IMAGES} 张图片",
                {"image_count": len(payload["images"])},
            )
        options = payload.get("options", {})
        if not isinstance(options, dict):
            raise AgentError("INVALID_REQUEST", "options 必须是 JSON 对象")
        unknown_options = sorted(set(options) - _OPTION_FIELDS)
        if unknown_options:
            raise AgentError(
                "INVALID_REQUEST",
                "options 包含未知字段",
                {"unknown_fields": unknown_options},
            )
        if options.get("if_exists", "fail") not in {"fail", "skip", "overwrite"}:
            raise AgentError("INVALID_REQUEST", "if_exists 必须是 fail、skip 或 overwrite")
        if options.get("on_image_error", "fail") != "fail":
            raise AgentError("INVALID_REQUEST", "当前仅支持 on_image_error=fail")
        for key in ("open", "use_settings"):
            if options.get(key, False) is not False:
                raise AgentError("INVALID_REQUEST", f"无头模式不支持 {key}=true")
        normalized = dict(payload)
        normalized["options"] = {
            "if_exists": options.get("if_exists", "fail"),
            "on_image_error": "fail",
        }
        return normalized

    @staticmethod
    def _validate_output(output: Path) -> None:
        if output.exists() and output.is_dir():
            raise AgentError("OUTPUT_NOT_WRITABLE", "输出路径是目录")
        parent = output.parent
        if not parent.is_dir():
            raise AgentError(
                "OUTPUT_NOT_WRITABLE", "输出目录不存在", {"output": str(output)}
            )

    @staticmethod
    def _map_error(error: Exception) -> tuple[str, str, dict]:
        if isinstance(error, ValueError):
            return "IMAGE_READ_FAILED", str(error), {}
        if isinstance(error, FileNotFoundError):
            return "TEMPLATE_MISSING", "所需模板不存在", {}
        if isinstance(error, (PermissionError, IsADirectoryError)):
            return "OUTPUT_NOT_WRITABLE", "无法写入输出文件，请检查路径或文件占用", {}
        if isinstance(error, OSError):
            return "OUTPUT_NOT_WRITABLE", "文档保存失败，请检查输出目录", {}
        return "INTERNAL_ERROR", "发生未预期错误，请查看诊断日志", {}

    @staticmethod
    def _success(
        job_id: str,
        output: Path,
        image_count: int,
        duplicates: int,
        warnings: list[str],
        started: float,
        **extra,
    ) -> dict:
        return {
            "ok": True,
            "job_id": job_id,
            "output": str(output),
            "image_count": image_count,
            "duplicates_skipped": duplicates,
            "warnings": warnings,
            "elapsed_ms": round((time.perf_counter() - started) * 1000, 3),
            "app_version": APP_VERSION,
            **extra,
        }
