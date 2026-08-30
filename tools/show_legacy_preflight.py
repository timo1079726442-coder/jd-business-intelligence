import json
d=json.load(open('data/legacy_excel_preflight.json',encoding='utf-8'))
for b in d:
 print('\n'+b['file'])
 for s in b['sheets']:
  print(s['name'],s['rows'],s['cols'])
