from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class TemplateIssue:
    template_number: int
    code: str
    message: str


@dataclass(frozen=True, slots=True)
class TemplateValidationResult:
    valid: bool
    checked_count: int
    issues: tuple[TemplateIssue, ...]
