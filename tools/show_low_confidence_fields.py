import json
from pathlib import Path
items=json.loads((Path(__file__).resolve().parents[1]/'data'/'legacy_mapping_preview.json').read_text(encoding='utf-8'))
for x in items:
 if x['source_sheet'] in ('快车订单','全站订单'):
  print(x['source_file'],x['source_sheet'],'target=',x['target_table'])
  print('matched=',[(m['excel_header'],m['target_column']) for m in x['matches']])
  print('unmatched=',x['unmatched_headers'])
