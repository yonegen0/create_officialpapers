"""User-facing application errors."""

class ConversionError(Exception):
    """Input, parsing or output error safe to display to users."""


class InputChangedError(ConversionError):
    """Input changed since preview."""
