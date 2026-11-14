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


def test_open_failures_keep_generated_document(tmp_path, monkeypatch):
    root, view = _window()
    try:
        from markdown_to_word.gui import controller, dialogs
        output = tmp_path / "既存.docx"
        output.write_bytes(b"existing")
        view.controller.last_output = output
        errors = []
        monkeypatch.setattr(controller.os, "startfile", lambda _: (_ for _ in ()).throw(OSError("起動できません")), raising=False)
        monkeypatch.setattr(dialogs, "error", lambda title, message, parent: errors.append(title))
        view.controller.open_word()
        view.controller.open_folder()
        assert len(errors) == 2
        assert output.read_bytes() == b"existing"
    finally:
        view.controller.close()


def test_preview_metrics_invalidate_without_losing_last_output(tmp_path):
    root, view = _window()
    try:
        view.input_mode.set("paste")
        view.input_text.insert("1.0", "# 見出し\n\n本文")
        root.update()
        view.controller.confirm()
        assert view.metric_values["headings"].get() == "1"
        assert view.metric_values["characters"].get() != "—"
        previous = tmp_path / "前回.docx"
        previous.write_bytes(b"previous")
        view.controller.last_output = previous
        view.show_success(previous)

        view.input_text.insert("end-1c", "\n追加")
        root.update()
        assert view.controller.preview is None
        assert all(metric.get() == "—" for metric in view.metric_values.values())
        assert "前回生成したWordファイル" in view.success.get()
        assert view.controller.last_output == previous

        view.controller.confirm()
        assert view.controller.preview is not None
        view.output_name.set("別名.docx")
        view.output_dir.set(str(tmp_path))
        assert view.controller.preview is not None
    finally:
        view.controller.close()


def test_busy_disables_editing_and_output_actions(tmp_path):
    root, view = _window()
    try:
        path = tmp_path / "前回.docx"
        path.write_bytes(b"previous")
        view.controller.last_output = path
        view.show_success(path)

        view.set_busy(True)
        assert str(view.name_entry.cget("state")) == "disabled"
        assert str(view.dir_entry.cget("state")) == "disabled"
        assert str(view.input_text.cget("state")) == "disabled"
        assert str(view.open_button.cget("state")) == "disabled"
        assert str(view.open_folder_button.cget("state")) == "disabled"
        assert "前回生成したWordファイル" in view.success.get()

        view.set_busy(False)
        assert str(view.input_text.cget("state")) == "normal"
        assert str(view.style_combo.cget("state")) == "readonly"
        assert str(view.open_button.cget("state")) == "normal"
    finally:
        view.controller.close()


def test_metrics_reset_but_previous_output_is_kept(tmp_path):
    root, view = _window()
    try:
        view.input_mode.set("paste")
        view.input_text.insert("1.0", "# 見出し\n\n- 項目")
        root.update()
        view.controller.confirm()
        assert view.metric_values["headings"].get() != "—"
        assert view.metric_values["list_items"].get() != "—"
        old_path = tmp_path / "前回.docx"
        view.controller.last_output = old_path
        view.show_success(old_path)
        view.input_text.insert("end", "\n追記")
        root.update()
        assert view.controller.preview is None
        assert all(value.get() == "—" for value in view.metric_values.values())
        assert "前回生成したWordファイル" in view.success.get()
        assert old_path.name in view.success.get()
    finally:
        view.controller.close()


def test_busy_disables_all_mutable_controls_and_output_actions(tmp_path):
    root, view = _window()
    try:
        view.controller.last_output = tmp_path / "前回.docx"
        view.set_busy(True)
        for widget in (
            view.file_button, view.folder_button, view.preview_button,
            view.generate_button, view.name_entry, view.dir_entry,
            view.open_button, view.open_folder_button, *view.mode_widgets,
        ):
            assert str(widget.cget("state")) == "disabled"
        assert str(view.style_combo.cget("state")) == "disabled"
        assert str(view.input_text.cget("state")) == "disabled"
        view.set_busy(False)
        assert str(view.style_combo.cget("state")) == "readonly"
        assert str(view.input_text.cget("state")) == "normal"
        assert str(view.open_button.cget("state")) == "normal"
    finally:
        view.controller.close()
