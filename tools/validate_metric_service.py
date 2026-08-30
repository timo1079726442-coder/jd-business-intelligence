import json
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from metrics.service import MetricService

result = MetricService().get_shop_overview("MIYO", "2026-08-20", "2026-08-26")
print(json.dumps(result, ensure_ascii=False, indent=2))
