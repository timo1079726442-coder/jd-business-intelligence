import json
d=json.load(open('data/legacy_excel_preflight.json',encoding='utf-8'))
keep={'订单列表','售后','推广自定义报表','快车订单','全站','全站-全店','全站订单','搜索','推荐','购物车'}
for b in d:
 for s in b['sheets']:
  if s['name'] in keep:
   print(b['file'],s['name'],s['headers'])
