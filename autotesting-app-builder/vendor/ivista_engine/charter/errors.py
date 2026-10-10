"""Stable, content-safe compiler errors."""

from __future__ import annotations


class CharterError(ValueError):
    def __init__(self, code: str, stage: str, message: str, details: dict | None = None):
        super().__init__(message)
        self.code = code
        self.stage = stage
        self.message = message
        self.details = details or {}

    def as_dict(self) -> dict:
        return {"errorCode": self.code, "stage": self.stage, "message": self.message, "details": self.details}
