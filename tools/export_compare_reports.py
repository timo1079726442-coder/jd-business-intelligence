import os,sys
from pathlib import Path

# Windows 默认控制台常为 GBK；项目日志含中文/emoji 时会触发编码异常。
# 统一切换为 UTF-8，避免导出完成后因日志输出失败而误报任务失败。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from main import run_business

for shop in ('FYA箱包旗舰店','MIYO箱包旗舰店','OTA箱包旗舰店'):
    os.environ['SHOP_ID']=shop
    for biz in ('京准通快车订单效果明细','京准通全站营销单品推广效果'):
        print('RUN',shop,biz)
        try:
            result=run_business(biz,start_date='2026-08-27',end_date='2026-08-27')
            print('OK',result)
        except Exception as exc:
            print('FAIL',shop,biz,type(exc).__name__,str(exc)[:300])
