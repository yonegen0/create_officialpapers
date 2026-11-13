import time
import tkinter as tk

import pytest


def _window():
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("GUI display unavailable")
    root.withdraw()
    from markdown_to_word.gui.view import TkView
    return root, TkView(root)


def test_input_modes_preserve_text_and_file_selection(tmp_path, monkeypatch):
    root, view = _window()
    try:
        source = tmp_path / "架空.md"
        source.write_text("# ファイル", encoding="utf-8")
        from markdown_to_word.gui import controller
        monkeypatch.setattr(controller.filedialog, "askopenfilename", lambda **kw: str(source))
        view.controller.choose_file()
        assert view.file_path.get() == str(source)
        assert view.output_name.get() == "架空.docx"
        assert view.output_dir.get() == str(tmp_path)
        view.input_text.insert("1.0", "# 貼り付け内容")
        view.input_mode.set("paste")
        assert view.input_text.get("1.0", "end-1c") == "# 貼り付け内容"
        view.input_mode.set("file")
        assert view.file_path.get() == str(source)
    finally:
        view.controller.close()


def test_gui_preview_generate_busy_and_success(tmp_path):
    root, view = _window()
    try:
        view.input_mode.set("paste")
        view.input_text.insert("1.0", "# 見出し\n\n本文です。")
        view.output_name.set("出力.docx")
        view.output_dir.set(str(tmp_path))
        root.update()
        view.controller.confirm()
        assert view.controller.preview is not None
        view.controller.generate()
        assert view.controller.busy
        assert str(view.input_text.cget("state")) == "disabled"
        for _ in range(200):
            root.update()
            if not view.controller.busy:
                break
            time.sleep(0.01)
        assert not view.controller.busy
        assert (tmp_path / "出力.docx").exists()
        assert view.controller.last_output is not None
        assert "完了" in view.status.get()
    finally:
        view.controller.close()
