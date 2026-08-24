from __future__ import annotations

import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_DIR / "src"))

from honyu_app.main import main


if __name__ == "__main__":
    raise SystemExit(main())
