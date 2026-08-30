import json
from pathlib import Path

root=Path(__file__).resolve().parents[1]
items=json.loads((root/'data'/'legacy_mapping_preview.json').read_text(encoding='utf-8'))
for item in items:
    print(item['source_file'], item['source_sheet'], item['target_table'], f"{item['matched_columns']}/{item['total_headers']}", item['historical_rows'])
