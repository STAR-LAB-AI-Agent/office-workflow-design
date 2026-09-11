# -*- coding: utf-8 -*-
"""
excel_summary.py — 读取 Excel，生成统计摘要 JSON（工作流第 1 步）
用法: python scripts/excel_summary.py --input test_data/销售数据.xlsx --output outputs/摘要.json
统一接口: 向屏幕打印一行 JSON {"status": ..., "outputs": {...}, "summary": ...}
"""
import argparse
import json
import sys
from pathlib import Path

import pandas as pd


def fail(msg):
    print(json.dumps({"status": "error", "error": msg}, ensure_ascii=False))
    sys.exit(1)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="读取 Excel 并生成统计摘要 JSON")
    parser.add_argument("--input", required=True, help="输入 Excel 文件路径")
    parser.add_argument("--output", required=True, help="输出摘要 JSON 文件路径")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        fail(f"输入文件不存在: {input_path}")

    try:
        df = pd.read_excel(input_path)
    except Exception as e:
        fail(f"Excel 读取失败: {e}")

    if df.empty:
        fail("Excel 内容为空")

    rows, cols = df.shape
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    text_cols = [c for c in df.columns if c not in numeric_cols]

    # 数值列统计
    numeric_stats = {}
    for c in numeric_cols:
        s = df[c].dropna()
        if len(s) == 0:
            continue
        numeric_stats[c] = {
            "sum": round(float(s.sum()), 2),
            "mean": round(float(s.mean()), 2),
            "max": round(float(s.max()), 2),
            "min": round(float(s.min()), 2),
        }

    # 分类列的分组汇总（只对取值种类 <=20 的列做，避免按订单号分组没意义）
    group_stats = {}
    if text_cols and numeric_cols:
        value_col = numeric_cols[0]
        for c in text_cols:
            if df[c].nunique() > 20:
                continue
            grouped = df.groupby(c)[value_col].sum().sort_values(ascending=False)
            group_stats[c] = [
                {"name": str(idx), "value": round(float(v), 2)}
                for idx, v in grouped.head(10).items()
            ]

    result_data = {
        "source_file": str(input_path),
        "rows": int(rows),
        "columns": int(cols),
        "column_names": [str(c) for c in df.columns],
        "numeric_stats": numeric_stats,
        "group_stats": group_stats,
    }

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(result_data, ensure_ascii=False, indent=2), encoding="utf-8")

    # 给智能体/用户看的一句话摘要
    parts = [f"共 {rows} 行 × {cols} 列"]
    for c, st in numeric_stats.items():
        parts.append(f"「{c}」合计 {st['sum']}，平均 {st['mean']}")
    for c, items in group_stats.items():
        if items:
            parts.append(f"按「{c}」汇总最高的是 {items[0]['name']}（{items[0]['value']}）")

    print(json.dumps({
        "status": "success",
        "input": str(input_path),
        "outputs": {"summary_json": str(output_path)},
        "summary": "；".join(parts) + f"。摘要已保存到 {output_path}",
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()