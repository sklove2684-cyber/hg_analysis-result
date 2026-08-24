from __future__ import annotations

from datetime import datetime
import os
from pathlib import Path
from time import perf_counter


_STARTED = perf_counter()


def mark(stage: str) -> None:
    elapsed = perf_counter() - _STARTED
    line = f"{datetime.now().isoformat(timespec='milliseconds')} +{elapsed:.3f}s {stage}"
    print(f"[STARTUP] {line}", flush=True)
    base = Path(os.getenv("LOCALAPPDATA") or Path.cwd()) / "HonyuAutomation"
    try:
        base.mkdir(parents=True, exist_ok=True)
        with (base / "startup.log").open("a", encoding="utf-8") as log:
            log.write(line + "\n")
    except OSError:
        pass
