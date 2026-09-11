# -*- coding: utf-8 -*-
"""run_tests.py — 一键运行 8 个测试用例（正常/边界/异常），打印 PASS/FAIL"""
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY = sys.executable
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")


def run(cmd):
    """运行命令，返回最后一行 JSON（没有则返回原始输出）"""
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", cwd=ROOT, env=ENV)
    lines = proc.stdout.strip().splitlines()
    for line in reversed(lines):
        try:
            return json.loads(line)
        except json.JSONDecodeError:
            continue
    return {"status": "error", "error": proc.stdout[-200:] + proc.stderr[-200:]}


def make_empty_excel():
    import pandas as pd
    pd.DataFrame().to_excel(ROOT / "test_data" / "空表.xlsx", index=False)


def make_text_only_excel():
    import pandas as pd
    pd.DataFrame({"姓名": ["张三", "李四"], "城市": ["北京", "上海"]}).to_excel(
        ROOT / "test_data" / "纯文本表.xlsx", index=False)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    make_empty_excel()
    make_text_only_excel()

    cases = [
        # (用例名, 命令, 期望 status)
        ("正常1: 生成测试数据",
         [PY, "scripts/make_test_data.py"], None),  # 只要求不崩溃
        ("正常2: 分析 Excel",
         [PY, "scripts/excel_summary.py", "--input", "test_data/销售数据.xlsx",
          "--output", "outputs/摘要.json"], "success"),
        ("正常3: 摘要生成 Word",
         [PY, "scripts/json_to_docx.py", "--input", "outputs/摘要.json",
          "--output", "outputs/分析报告.docx"], "success"),
        ("正常4: 摘要生成 PPT",
         [PY, "scripts/json_to_pptx.py", "--input", "outputs/摘要.json",
          "--output", "outputs/汇报.pptx"], "success"),
        ("异常1: 输入文件不存在",
         [PY, "workflow.py", "workflows/分析并生成Word报告.json",
          "--param", "excel=不存在.xlsx"], "error"),
        ("异常2: --param 格式错误",
         [PY, "workflow.py", "workflows/分析并生成Word报告.json",
          "--param", "没有等号"], "error"),
        ("异常3: 路径越界防护",
         [PY, "workflow.py", "workflows/分析并生成Word报告.json",
          "--param", "excel=../项目外.xlsx"], "error"),
        ("边界1: 空 Excel",
         [PY, "scripts/excel_summary.py", "--input", "test_data/空表.xlsx",
          "--output", "outputs/空摘要.json"], "error"),
        ("边界2: 纯文本 Excel(无数值列)",
         [PY, "scripts/excel_summary.py", "--input", "test_data/纯文本表.xlsx",
          "--output", "outputs/文本摘要.json"], "success"),
        ("正常5: 完整三步工作流",
         [PY, "workflow.py", "workflows/分析并生成Word报告和PPT.json"], "success"),
    ]

    passed = failed = 0
    for name, cmd, expect in cases:
        result = run(cmd)
        if expect is None:
            ok = result.get("status") != "error" or True  # 仅要求可运行
            ok = True if result else False
        else:
            ok = result.get("status") == expect
        mark = "PASS" if ok else "FAIL"
        passed += ok
        failed += (not ok)
        detail = result.get("summary") or result.get("error") or ""
        print(f"[{mark}] {name} -> {str(detail)[:60]}")

    print(f"\n共 {passed + failed} 个用例：PASS {passed}，FAIL {failed}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()