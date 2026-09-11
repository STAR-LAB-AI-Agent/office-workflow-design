# -*- coding: utf-8 -*-
"""make_test_data.py — 生成虚构的测试销售数据 Excel（仅供课程测试，无真实信息）"""
from pathlib import Path
import random
import pandas as pd

random.seed(42)
regions = ["华东", "华北", "华南", "西南"]
products = ["笔记本电脑", "手机", "平板", "耳机"]

rows = []
for i in range(1, 121):
    rows.append({
        "订单号": f"SO{i:04d}",
        "日期": f"2025-{random.randint(1, 6):02d}-{random.randint(1, 28):02d}",
        "地区": random.choice(regions),
        "产品": random.choice(products),
        "销售额": round(random.uniform(500, 20000), 2),
        "数量": random.randint(1, 20),
    })

out = Path("test_data/销售数据.xlsx")
out.parent.mkdir(parents=True, exist_ok=True)
pd.DataFrame(rows).to_excel(out, index=False)
print(f"已生成测试数据: {out}（120 行，虚构数据）")