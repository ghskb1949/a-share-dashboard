from __future__ import annotations

import json
import re
from pathlib import Path


html = Path("output/dashboard.html").read_text(encoding="utf-8")

payload_match = re.search(
    r'<script id="payload" type="application/json">(.*?)</script>',
    html,
    re.S,
)
if not payload_match:
    raise SystemExit("payload script not found")

payload = json.loads(payload_match.group(1))
required = ["overview", "maRanges", "daily", "marketSectors", "review"]
missing = [key for key in required if not payload.get(key)]
if missing:
    raise SystemExit(f"missing payload sections: {missing}")

checks = {
    "trendCanvas": "trendCanvas" in html,
    "stock-card": "stock-card" in html,
    "MA panel": "均线高低点" in html,
    "daily table": "原始日线" in html,
    "review": "三轮复盘" in html,
    "stock names": all(name in html for name in ["新洁能", "三安光电", "蔚蓝锂芯", "中天科技"]),
}
failed = [name for name, ok in checks.items() if not ok]
if failed:
    raise SystemExit(f"failed checks: {failed}")

print("DASHBOARD_VERIFY_OK")
print("overview", len(payload["overview"]))
print("maRanges", len(payload["maRanges"]))
print("daily", len(payload["daily"]))
