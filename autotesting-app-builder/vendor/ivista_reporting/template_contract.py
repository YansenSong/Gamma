"""Shared exception for versioned DOCX template contract failures."""


class TemplateContractError(RuntimeError):
    """Raised when a machine report template violates its binding contract."""
