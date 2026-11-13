"""Application entry point kept separate from the GUI for testability."""
from .service import prepare, convert, read_file

__all__ = ["prepare", "convert", "read_file"]
