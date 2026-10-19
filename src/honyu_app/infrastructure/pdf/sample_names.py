from __future__ import annotations

import re


def is_recovery_blank(value: str) -> bool:
    """Recognize recovery blanks regardless of separators or English case."""
    return re.sub(r"[\s_-]+", "", value).casefold() in {"회수율b", "회수율blank"}


def canonical_recovery_sample_name(value: str) -> str:
    return "회수율-B" if is_recovery_blank(value) else value
