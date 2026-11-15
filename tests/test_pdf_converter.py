"""Mocked Word automation tests; never start real Microsoft Word in CI."""
from __future__ import annotations

from pathlib import Path
from types import ModuleType, SimpleNamespace
import sys

from markdown_to_word import pdf_converter
from markdown_to_word.pdf_converter import WordConnectionError, WordSession, convert_pdf_files

PDF = b"%PDF-1.7\n1 0 obj\n<<>>\nendobj\n%%EOF\n"


class FakeSession:
    def __init__(self, failures=None):
        self.failures = failures or {}
        self.opened = []
        self.closed = False
        self.alive = True
    def __enter__(self):
        return self
    def __exit__(self, *args):
        self.closed = True
    def export(self, source, temp):
        self.opened.append(source)
        if source.name in self.failures:
            error = self.failures[source.name]
            if isinstance(error, WordConnectionError):
                self.alive = False
            raise error
        temp.write_bytes(PDF)
    def is_alive(self):
        return self.alive


def test_single_and_batch_keep_sources_and_number_outputs(tmp_path):
    paths = tuple(tmp_path / f"会社{i}.docx" for i in range(2))
    for source in paths:
        source.write_bytes(b"original Word")
    (tmp_path / "会社0.pdf").write_bytes(b"old PDF")
    instance = FakeSession()
    updates = []
    result = convert_pdf_files(
        paths, tmp_path, session_factory=lambda: instance,
        progress=lambda *args: updates.append(args),
    )
    assert result.processed == result.total == 2
    assert len(result.saved) == 2
    assert len(result.failures) == 0
    assert result.saved[0].name == "会社0 (1).pdf"
    assert (tmp_path / "会社0.pdf").read_bytes() == b"old PDF"
    assert all(path.read_bytes() == b"original Word" for path in paths)
    assert len(updates) == 2
    assert instance.closed
    assert not list(tmp_path.glob(".mtw-pdf-*.pdf"))


def test_per_file_errors_continue_and_fatal_word_failure_stops(tmp_path):
    files = tuple(tmp_path / f"{name}.docx" for name in ("ok", "bad", "later"))
    for file in files:
        file.write_bytes(b"Word")
    instance = FakeSession({"bad.docx": ValueError("password protected")})
    result = convert_pdf_files(files, tmp_path, session_factory=lambda: instance)
    assert result.processed == 3
    assert len(result.saved) == 2 and len(result.failures) == 1
    assert result.stopped_reason is None

    broken = FakeSession({"bad.docx": WordConnectionError("disconnected")})
    result = convert_pdf_files(files, tmp_path, session_factory=lambda: broken)
    assert result.processed == 2
    assert len(result.failures) == 1
    assert result.stopped_reason is not None
    assert files[2] not in broken.opened
    assert broken.closed


def test_missing_source_is_one_file_failure(tmp_path):
    missing = tmp_path / "missing.docx"
    good = tmp_path / "good.docx"
    good.write_bytes(b"untouched")
    result = convert_pdf_files((missing, good), tmp_path, session_factory=FakeSession)
    assert result.processed == 2
    assert len(result.saved) == len(result.failures) == 1
    assert good.read_bytes() == b"untouched"


def test_word_session_initializes_and_releases_com_on_same_thread(monkeypatch, tmp_path):
    calls = []
    com = ModuleType("pythoncom")
    com.CoInitialize = lambda: calls.append("initialize")
    com.CoUninitialize = lambda: calls.append("uninitialize")
    win32com = ModuleType("win32com")
    win32com.__path__ = []
    client = ModuleType("win32com.client")
    class WordDocument:
        def ExportAsFixedFormat(self, **kwargs):
            calls.append(("export", kwargs))
            Path(kwargs["OutputFileName"]).write_bytes(PDF)
        def Close(self, **kwargs):
            calls.append(("close", kwargs))
    class WordApp:
        Version = "16.0"
        Documents = SimpleNamespace(Open=lambda **kwargs: (calls.append(("open", kwargs)), WordDocument())[1])
        def Quit(self, **kwargs):
            calls.append(("quit", kwargs))
    app = WordApp()
    client.DispatchEx = lambda name: (calls.append(("dispatch", name)), app)[1]
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", com)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(pdf_converter, "os", SimpleNamespace(name="nt"))
    with WordSession() as session:
        session.export(tmp_path / "input.docx", tmp_path / "output.pdf")
    assert calls[0] == "initialize" and calls[-1] == "uninitialize"
    assert calls[1] == ("dispatch", "Word.Application")
    assert calls[2][1]["ReadOnly"] is True
    assert calls[2][1]["UpdateLinks"] == 0
    assert calls[3][1]["ExportFormat"] == 17
    assert calls[3][1]["OpenAfterExport"] is False
    assert calls[3][1]["Item"] == 0
    assert calls[4] == ("close", {"SaveChanges": 0})
    assert calls[5] == ("quit", {"SaveChanges": 0})
    assert app.AutomationSecurity == 3



def test_temp_cleanup_warning_does_not_fail_saved_pdfs_or_stop_batch(tmp_path, monkeypatch):
    files = tuple(tmp_path / f"sample{i}.docx" for i in range(2))
    for source in files:
        source.write_bytes(b"untouched Word")
    real_unlink = Path.unlink
    session = FakeSession()

    def deny_temp_cleanup(path, *args, **kwargs):
        if path.name.startswith(".mtw-pdf-"):
            raise PermissionError("locked by Windows")
        return real_unlink(path, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", deny_temp_cleanup)
    result = convert_pdf_files(files, tmp_path, session_factory=lambda: session)
    assert result.processed == result.total == 2
    assert len(result.saved) == 2 and len(result.failures) == 0
    assert len(result.warnings) == 2
    assert all("一時PDF" in warning for _, warning in result.warnings)
    assert result.stopped_reason is None
    assert session.opened == list(files) and session.closed
    assert all(path.read_bytes() == b"untouched Word" for path in files)
    assert all(path.read_bytes() == PDF for path in result.saved)


def test_startup_failure_is_not_counted_as_file_failure(tmp_path):
    from markdown_to_word.gui.pdf_view import PdfTab

    file = tmp_path / "not_processed.docx"
    file.write_bytes(b"original")

    class CannotStart:
        def __enter__(self):
            raise WordConnectionError("Microsoft Wordを起動できません。")

    result = convert_pdf_files((file,), tmp_path, session_factory=CannotStart)
    assert result.processed == 0 and result.total == 1
    assert not result.saved and not result.failures and not result.warnings
    assert result.stopped_reason is not None
    output = []
    view = SimpleNamespace(
        result=SimpleNamespace(set=output.append),
        _result_actions=lambda: None,
    )
    PdfTab.show_result(view, result)
    message = output[0]
    assert "PDF変換を開始できませんでした" in message
    assert "デスクトップ版Microsoft Word" in message
    assert "未処理：1件" in message
    assert "0件失敗" not in message
    assert file.read_bytes() == b"original"


def test_stopped_batch_has_consistent_processed_and_unprocessed_counts(tmp_path):
    from markdown_to_word.gui.pdf_view import PdfTab

    files = tuple(tmp_path / f"{name}.docx" for name in ("ok", "fatal", "later"))
    for file in files:
        file.write_bytes(b"original")
    session = FakeSession({"fatal.docx": WordConnectionError("connection lost")})
    result = convert_pdf_files(files, tmp_path, session_factory=lambda: session)
    assert result.processed == 2 and result.total == 3
    assert len(result.saved) == 1 and len(result.failures) == 1
    assert len(result.warnings) == 0
    output = []
    view = SimpleNamespace(
        result=SimpleNamespace(set=output.append),
        _result_actions=lambda: None,
    )
    PdfTab.show_result(view, result)
    assert "3件中1件成功・1件失敗" in output[0]
    assert "未処理：1件" in output[0]
    assert files[2].read_bytes() == b"original"
