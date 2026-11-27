from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

from PIL import Image

PROJECT = Path(__file__).resolve().parents[1]
PYTHON = sys.executable


def request(tmp_path: Path, name: str = "out.docx") -> dict:
    image = tmp_path / "photo.jpg"
    if not image.exists():
        Image.new("RGB", (64, 48), (40, 80, 120)).save(image)
    return {
        "schema_version": 1,
        "title": "自动化活动",
        "location": "测试地点",
        "activity_date": "2026-11-27",
        "images": [str(image)],
        "output": str(tmp_path / name),
    }


def invoke(arguments: list[str], input_text: str | None = None):
    return subprocess.run(
        [PYTHON, str(PROJECT / "main.py"), *arguments, "--json"],
        input=input_text,
        text=True,
        encoding="utf-8",
        capture_output=True,
        cwd=PROJECT,
        timeout=30,
    )


def test_real_generate_has_clean_json_and_valid_docx(tmp_path: Path) -> None:
    output = tmp_path / "result.docx"
    payload = request(tmp_path, output.name)
    completed = invoke(["--generate", "-"], json.dumps(payload, ensure_ascii=False))

    assert completed.returncode == 0, completed.stderr
    result = json.loads(completed.stdout)
    assert result["ok"] is True
    assert result["image_count"] == 1
    assert result["output"] == str(output)
    assert result["job_id"]
    assert "PyQt6" not in completed.stderr
    with ZipFile(output) as archive:
        assert "word/document.xml" in archive.namelist()


def test_real_generate_does_not_import_qt(tmp_path: Path) -> None:
    payload = request(tmp_path)
    wrapper = tmp_path / "no_qt.py"
    wrapper.write_text(
        "import runpy,sys;sys.path.insert(0,'.');\n"
        "sys.argv=['main.py','--generate','-','--json']\n"
        "try:\n"
        " runpy.run_path('main.py',run_name='__main__')\n"
        "except SystemExit as e:\n"
        " if e.code: raise\n"
        " if any(m == 'PyQt6' or m.startswith('PyQt6.') for m in sys.modules):\n"
        "  raise RuntimeError('Qt loaded')\n",
        encoding="utf-8",
    )
    completed = subprocess.run(
        [PYTHON, str(wrapper)],
        input=json.dumps(payload), text=True, encoding="utf-8",
        capture_output=True, cwd=PROJECT, timeout=30,
    )
    assert completed.returncode == 0


def test_existing_output_defaults_to_fail_then_skip(tmp_path: Path) -> None:
    payload = request(tmp_path)
    Path(payload["output"]).write_bytes(b"existing")
    completed = invoke(["--generate", "-"], json.dumps(payload))
    assert completed.returncode == 4
    assert json.loads(completed.stdout)["error"]["code"] == "OUTPUT_EXISTS"
    assert Path(payload["output"]).read_bytes() == b"existing"

    payload["options"] = {"if_exists": "skip"}
    completed = invoke(["--generate", "-"], json.dumps(payload))
    assert completed.returncode == 0
    assert json.loads(completed.stdout)["skipped"] is True
    assert Path(payload["output"]).read_bytes() == b"existing"


def test_invalid_and_missing_images_have_structured_errors(tmp_path: Path) -> None:
    payload = request(tmp_path)
    payload["images"] = [str(tmp_path / "absent.jpg")]
    completed = invoke(["--generate", "-"], json.dumps(payload))
    assert completed.returncode == 2
    assert json.loads(completed.stdout)["error"]["code"] == "IMAGE_NOT_FOUND"

    payload["images"] = [str(tmp_path / "bad.txt")]
    Path(payload["images"][0]).write_text("x")
    completed = invoke(["--generate", "-"], json.dumps(payload))
    assert json.loads(completed.stdout)["error"]["code"] == "UNSUPPORTED_FORMAT"

    payload["unexpected"] = True
    completed = invoke(["--validate", "-"], json.dumps(payload))
    assert completed.returncode == 2
    error = json.loads(completed.stdout)["error"]
    assert error["code"] == "INVALID_REQUEST"
    assert error["details"]["unknown_fields"] == ["unexpected"]


def test_corrupt_image_does_not_leave_partial_docx(tmp_path: Path) -> None:
    payload = request(tmp_path)
    payload["images"] = [str(tmp_path / "corrupt.jpg")]
    Path(payload["images"][0]).write_bytes(b"not an image")
    completed = invoke(["--generate", "-"], json.dumps(payload))
    assert completed.returncode == 2
    assert json.loads(completed.stdout)["error"]["code"] == "IMAGE_READ_FAILED"
    assert not Path(payload["output"]).exists()


def test_validate_capabilities_and_schema(tmp_path: Path) -> None:
    completed = invoke(["--validate", "-"], json.dumps(request(tmp_path)))
    assert completed.returncode == 0
    assert json.loads(completed.stdout)["validated"] is True
    assert not Path(request(tmp_path)["output"]).exists()

    capabilities = json.loads(invoke(["--capabilities"]).stdout)
    assert capabilities["max_images"] == 30
    assert capabilities["templates"]["missing"] == []
    schema = json.loads(invoke(["--schema"]).stdout)
    assert schema["additionalProperties"] is False
    assert schema["properties"]["schema_version"]["const"] == 1


def test_batch_continues_after_bad_line(tmp_path: Path) -> None:
    valid = request(tmp_path)
    batch = tmp_path / "jobs.jsonl"
    batch.write_text(
        "\n".join(["# comment", "{bad", json.dumps(valid, ensure_ascii=False)]),
        encoding="utf-8",
    )
    completed = invoke(["--batch", str(batch)])
    lines = [json.loads(line) for line in completed.stdout.splitlines()]
    assert [line["ok"] for line in lines] == [False, True]
    assert lines[0]["line"] == 2
    assert Path(valid["output"]).exists()


def test_persistent_server_recovers_and_correlates_ids(tmp_path: Path) -> None:
    valid = request(tmp_path)
    requests = "\n".join([
        "{bad",
        json.dumps({"id": "first", "request": valid}, ensure_ascii=False),
        json.dumps({"id": 2, "request": {"schema_version": 1}}),
    ]) + "\n"
    completed = invoke(["--serve"], requests)
    lines = [json.loads(line) for line in completed.stdout.splitlines()]
    assert [line["id"] for line in lines] == [None, "first", 2]
    assert [line["result"]["ok"] for line in lines] == [False, True, False]
