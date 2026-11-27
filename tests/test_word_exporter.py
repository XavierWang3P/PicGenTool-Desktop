from __future__ import annotations

from pathlib import Path
import threading
import time
from unittest.mock import Mock

import pytest

from models.document_layout import WordExportPayload
from models.word_export_session import PreparedWordExport
from services.cancellation import CancellationToken, GenerationCancelled
from services.word_exporter import WordExporter


def _payload(tmp_path: Path) -> WordExportPayload:
    return WordExportPayload(
        template_path=tmp_path / "template.docx",
        text_context={"title": "test"},
        image_paths=[tmp_path / "image.jpg"],
        output_path=tmp_path / "result.docx",
    )


def _mock_docx(monkeypatch: pytest.MonkeyPatch, doc: Mock) -> None:
    import docxtpl

    monkeypatch.setattr(docxtpl, "DocxTemplate", Mock(return_value=doc))
    monkeypatch.setattr(docxtpl, "InlineImage", Mock(return_value="inline-image"))


def test_export_saves_to_temporary_file_then_replaces_target(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    payload.output_path.write_bytes(b"old")
    doc = Mock()
    saved_paths: list[Path] = []

    def save(path: str) -> None:
        saved_path = Path(path)
        saved_paths.append(saved_path)
        saved_path.write_bytes(b"new")

    doc.save.side_effect = save
    _mock_docx(monkeypatch, doc)

    result = WordExporter().export(payload)

    assert result == payload.output_path
    assert payload.output_path.read_bytes() == b"new"
    assert len(saved_paths) == 1
    assert saved_paths[0] != payload.output_path
    assert saved_paths[0].parent == payload.output_path.parent
    assert saved_paths[0].suffix == ".docx"
    assert not saved_paths[0].exists()


def test_export_failure_preserves_existing_target_and_removes_temporary_file(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    payload.output_path.write_bytes(b"old")
    doc = Mock()
    saved_paths: list[Path] = []

    def fail_save(path: str) -> None:
        saved_path = Path(path)
        saved_paths.append(saved_path)
        saved_path.write_bytes(b"partial")
        raise OSError("disk full")

    doc.save.side_effect = fail_save
    _mock_docx(monkeypatch, doc)

    with pytest.raises(IOError, match="文档生成失败"):
        WordExporter().export(payload)

    assert payload.output_path.read_bytes() == b"old"
    assert saved_paths and all(not path.exists() for path in saved_paths)


def test_export_cancelled_after_render_does_not_save_or_replace(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    payload.output_path.write_bytes(b"old")
    token = CancellationToken()
    doc = Mock()
    doc.render.side_effect = lambda _context: token.cancel()
    _mock_docx(monkeypatch, doc)

    with pytest.raises(GenerationCancelled):
        WordExporter().export(payload, cancellation=token)

    doc.save.assert_not_called()
    assert payload.output_path.read_bytes() == b"old"
    assert list(tmp_path.glob("*.docx")) == [payload.output_path]


def test_export_cancelled_during_save_preserves_target_and_cleans_temp(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    payload.output_path.write_bytes(b"old")
    token = CancellationToken()
    doc = Mock()

    def save(path: str) -> None:
        Path(path).write_bytes(b"new")
        token.cancel()

    doc.save.side_effect = save
    _mock_docx(monkeypatch, doc)

    with pytest.raises(GenerationCancelled):
        WordExporter().export(payload, cancellation=token)

    assert payload.output_path.read_bytes() == b"old"
    assert list(tmp_path.glob("*.docx")) == [payload.output_path]


def test_export_commit_blocks_late_cancel_and_reports_success(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    payload.output_path.write_bytes(b"old")
    token = CancellationToken()
    doc = Mock()
    doc.save.side_effect = lambda path: Path(path).write_bytes(b"new")
    _mock_docx(monkeypatch, doc)
    replace_started = threading.Event()
    allow_replace = threading.Event()
    original_replace = Path.replace

    def blocking_replace(path: Path, target: Path) -> Path:
        replace_started.set()
        assert allow_replace.wait(timeout=1)
        return original_replace(path, target)

    monkeypatch.setattr(Path, "replace", blocking_replace)
    results: list[Path] = []
    errors: list[BaseException] = []

    def export() -> None:
        try:
            results.append(WordExporter().export(payload, cancellation=token))
        except BaseException as exc:
            errors.append(exc)

    export_thread = threading.Thread(target=export)
    export_thread.start()
    assert replace_started.wait(timeout=1)
    cancel_thread = threading.Thread(target=token.cancel)
    cancel_thread.start()
    time.sleep(0.02)
    cancel_was_blocked = cancel_thread.is_alive()
    allow_replace.set()
    export_thread.join(timeout=1)
    cancel_thread.join(timeout=1)

    assert cancel_was_blocked
    assert not errors
    assert results == [payload.output_path]
    assert payload.output_path.read_bytes() == b"new"
    assert not token.is_cancelled()


def test_prepared_word_export_is_consumed_on_creator_thread(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    doc = Mock()
    doc.save.side_effect = lambda path: Path(path).write_bytes(b"new")
    _mock_docx(monkeypatch, doc)
    exporter = WordExporter()

    prepared = exporter.prepare(payload)
    result = exporter.export_prepared(prepared, payload.image_paths)

    assert isinstance(prepared, PreparedWordExport)
    assert prepared.creator_thread_id == threading.get_ident()
    assert result == payload.output_path


def test_prepared_word_export_rejects_cross_thread_consumption(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    doc = Mock()
    _mock_docx(monkeypatch, doc)
    exporter = WordExporter()
    prepared = exporter.prepare(payload)
    errors: list[BaseException] = []

    thread = threading.Thread(
        target=lambda: _capture_error(
            errors,
            lambda: exporter.export_prepared(prepared, payload.image_paths),
        )
    )
    thread.start()
    thread.join(timeout=1)

    assert len(errors) == 1
    assert isinstance(errors[0], RuntimeError)
    assert "创建线程" in str(errors[0])
    doc.render.assert_not_called()


def _capture_error(errors: list[BaseException], action) -> None:
    try:
        action()
    except BaseException as exc:
        errors.append(exc)


def test_word_open_error_does_not_expose_private_template_directory(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    payload = _payload(tmp_path)
    private_parent = tmp_path / "private-template"
    payload.template_path = private_parent / "Template-1.docx"
    _mock_docx(monkeypatch, Mock())
    import docxtpl
    docxtpl.DocxTemplate.side_effect = OSError(f"cannot open {private_parent}")

    with pytest.raises(IOError) as error:
        WordExporter().prepare(payload)

    assert str(private_parent) not in str(error.value)
    assert "模板文件打开失败" in str(error.value)


def test_word_save_error_and_cleanup_log_do_not_expose_private_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    payload = _payload(tmp_path)
    private_parent = tmp_path / "private-output"
    private_parent.mkdir()
    payload.output_path = private_parent / "result.docx"
    doc = Mock()
    saved_paths: list[Path] = []

    def fail_save(path: str) -> None:
        saved_paths.append(Path(path))
        raise OSError(f"failed for {path}")

    doc.save.side_effect = fail_save
    _mock_docx(monkeypatch, doc)
    monkeypatch.setattr(Path, "unlink", Mock(side_effect=PermissionError(str(private_parent))))

    with caplog.at_level("WARNING", logger="pgt.word"):
        with pytest.raises(IOError) as error:
            WordExporter().export(payload)

    assert str(private_parent) not in str(error.value)
    assert str(private_parent) not in caplog.text
    assert "文档生成失败" in str(error.value)
    assert saved_paths
