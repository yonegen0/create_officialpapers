# -*- mode: python ; coding: utf-8 -*-
# Build ONLY on Windows 11 x64 with CPython 3.13.
a = Analysis(
    ["markdown_to_word/gui_entry.py"],
    pathex=["."],
    binaries=[],
    datas=[
        ("markdown_to_word/assets/MarkdownToWord.ico", "markdown_to_word/assets"),
        ("markdown_to_word/assets/resume_template.docx", "markdown_to_word/assets"),
    ],
    hiddenimports=["markdown_it.rules_block.table"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "pypdf", "cryptography", "matplotlib"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="MarkdownToWord", debug=False,
          bootloader_ignore_signals=False, strip=False, upx=False, console=False,
          icon="markdown_to_word/assets/MarkdownToWord.ico")
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="MarkdownToWord")
