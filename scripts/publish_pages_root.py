from __future__ import annotations

import shutil
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]
OUTPUT_DIR = BASE_DIR / "output"


def latest_report() -> Path:
    reports = sorted(OUTPUT_DIR.glob("codex_daily_observation_*.xlsx"), reverse=True)
    if not reports:
        raise FileNotFoundError("No codex_daily_observation_*.xlsx report found in output.")
    return reports[0]


def main() -> None:
    dashboard = OUTPUT_DIR / "dashboard.html"
    if not dashboard.exists():
        raise FileNotFoundError(f"Missing dashboard: {dashboard}")

    report = latest_report()
    shutil.copy2(dashboard, BASE_DIR / "index.html")
    shutil.copy2(report, BASE_DIR / report.name)
    shutil.copy2(report, BASE_DIR / "latest_report.xlsx")
    (BASE_DIR / ".nojekyll").write_text("", encoding="utf-8")
    print(BASE_DIR / "index.html")
    print(BASE_DIR / report.name)
    print(BASE_DIR / "latest_report.xlsx")


if __name__ == "__main__":
    main()
