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
        monkeypatch.setattr(controller.filedialog, "askopenfilenames", lambda **kw: (str(source),))
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
        # This test covers GUI save flow, independent of the separate resume asset.
        view.style.set("標準文書")
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



def test_gui_batch_selection_preview_generation_and_paste_compatibility(tmp_path, monkeypatch):
    root, view = _window()
    try:
        from markdown_to_word.gui import controller
        paths = [tmp_path / f"企業{i}.md" for i in range(3)]
        for index, path in enumerate(paths):
            path.write_text(f"# 会社{index}\\n\\n本文", encoding="utf-8")
        view.style.set("標準文書")
        files = tuple(str(path) for path in paths)
        monkeypatch.setattr(controller.filedialog, "askopenfilenames", lambda **kw: files)
        view.controller.choose_file()
        assert len(view.controller.selected_files) == 3
        assert view.file_count.get() == "3件選択中"
        assert view.output_name.get() == "元ファイル名を使用"
        assert str(view.name_entry.cget("state")) == "disabled"

        # Cancelling the dialog must leave the selection unchanged.
        monkeypatch.setattr(controller.filedialog, "askopenfilenames", lambda **kw: ())
        view.controller.choose_file()
        assert len(view.controller.selected_files) == 3

        view.input_text.insert("1.0", "# 貼り付けた文章")
        view.input_mode.set("paste")
        assert view.input_text.get("1.0", "end-1c") == "# 貼り付けた文章"
        assert str(view.name_entry.cget("state")) == "normal"
        view.input_mode.set("file")
        assert str(view.name_entry.cget("state")) == "disabled"

        view.controller.confirm()
        assert view.controller.busy
        assert str(view.file_button.cget("state")) == "disabled"
        for _ in range(300):
            root.update()
            if not view.controller.busy:
                break
            time.sleep(0.01)
        assert not view.controller.busy
        assert view.metric_values["headings"].get() == "3"  # batch '生成可能' tile

        view.controller.generate()
        assert view.controller.busy
        assert str(view.dir_entry.cget("state")) == "disabled"
        assert str(view.exit_button.cget("state")) == "disabled"
        for _ in range(400):
            root.update()
            if not view.controller.busy:
                break
            time.sleep(0.01)
        assert not view.controller.busy
        assert len(list(tmp_path.glob("*.docx"))) == 3
        assert view.controller.last_batch_result is not None
        assert len(view.controller.last_batch_result.saved) == 3
        assert str(view.open_button.cget("state")) == "disabled"
        assert str(view.open_folder_button.cget("state")) == "normal"
        assert "3件成功" in view.success.get()

        # Switching back to a single file restores the editable filename.
        monkeypatch.setattr(controller.filedialog, "askopenfilenames", lambda **kw: (str(paths[0]),))
        view.controller.choose_file()
        assert view.output_name.get() == "企業0.docx"
        assert str(view.name_entry.cget("state")) == "normal"
    finally:
        view.controller.close()
