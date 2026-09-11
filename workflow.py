# -*- coding: utf-8 -*-
"""
workflow.py — 办公任务工作流引擎
按 JSON 流程定义依次执行各步骤脚本，支持步骤间结果传递。

用法:
  python workflow.py workflows\分析并生成Word报告和PPT.json
  python workflow.py workflows\分析并生成Word报告和PPT.json --param excel=test_data\销售数据.xlsx

流程定义格式:
{
  "name": "流程名",
  "params": {"excel": "默认输入.xlsx"},
  "steps": [
    {"id": "summary", "script": "scripts/excel_summary.py",
     "args": {"input": "{{params.excel}}", "output": "outputs/摘要.json"}},
    {"id": "docx", "script": "scripts/json_to_docx.py",
     "args": {"input": "{{steps.summary.outputs.summary_json}}", "output": "outputs/报告.docx"}}
  ]
}
"""
import argparse
import json
import logging
import os
import re
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
LOG_DIR = PROJECT_ROOT / "logs"
PLACEHOLDER = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")


def setup_logging():
    LOG_DIR.mkdir(exist_ok=True)
    logging.basicConfig(
        filename=LOG_DIR / "workflow.log",
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        encoding="utf-8",
    )


def resolve(value, context):
    """把 {{steps.xxx.outputs.yyy}} 占位符替换为前面步骤的真实输出"""
    if not isinstance(value, str):
        return value

    def repl(match):
        parts = match.group(1).strip().split(".")
        node = context
        for part in parts:
            if not isinstance(node, dict) or part not in node:
                raise KeyError(f"无法解析占位符: {match.group(0)}")
            node = node[part]
        return str(node)

    return PLACEHOLDER.sub(repl, value)


def check_path_safe(path_str):
    """最小权限: 只允许读写项目目录内的文件"""
    p = (PROJECT_ROOT / path_str).resolve()
    if p != PROJECT_ROOT and PROJECT_ROOT not in p.parents:
        raise PermissionError(f"路径越界，已拒绝: {path_str}（只允许项目目录内）")
    return path_str


def run_step(step, context):
    script = step["script"]
    args = {k: resolve(v, context) for k, v in step.get("args", {}).items()}
    for v in args.values():
        check_path_safe(v)

    cmd = [sys.executable, str(PROJECT_ROOT / script)]
    for k, v in args.items():
        cmd += [f"--{k}", v]

    logging.info(f"执行步骤 {step['id']}: {' '.join(cmd)}")
    env = dict(os.environ, PYTHONIOENCODING="utf-8")  # 保证中文不乱码
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace",
                          cwd=PROJECT_ROOT, env=env)

    lines = proc.stdout.strip().splitlines()
    last_line = lines[-1] if lines else ""
    try:
        result = json.loads(last_line)
    except json.JSONDecodeError:
        result = {"status": "error",
                  "error": f"脚本输出不是 JSON: {last_line or proc.stderr[-200:]}"}
    logging.info(f"步骤 {step['id']} 结果: {json.dumps(result, ensure_ascii=False)}")
    return result


def main():
    sys.stdout.reconfigure(encoding="utf-8")  # 保证中文输出不乱码
    parser = argparse.ArgumentParser(description="办公任务工作流引擎")
    parser.add_argument("workflow", help="流程定义 JSON 文件路径")
    parser.add_argument("--param", action="append", default=[],
                        help="覆盖参数，格式 key=value，可多次使用")
    args = parser.parse_args()

    setup_logging()
    wf_path = (PROJECT_ROOT / args.workflow).resolve()
    if not wf_path.exists():
        print(json.dumps({"status": "error", "error": f"流程文件不存在: {args.workflow}"},
                         ensure_ascii=False))
        sys.exit(1)

    wf = json.loads(wf_path.read_text(encoding="utf-8"))
    params = dict(wf.get("params", {}))
    for kv in args.param:
        if "=" not in kv:
            print(json.dumps({"status": "error", "error": f"--param 格式错误: {kv}"},
                             ensure_ascii=False))
            sys.exit(1)
        k, v = kv.split("=", 1)
        params[k] = v

    context = {"params": params, "steps": {}}
    steps = wf.get("steps", [])
    logging.info(f"===== 开始工作流: {wf.get('name')} =====")
    print(f"[工作流] 开始: {wf.get('name')}（共 {len(steps)} 步）")

    for i, step in enumerate(steps, 1):
        print(f"[工作流] 第 {i}/{len(steps)} 步: {step['id']} ...")
        try:
            result = run_step(step, context)
        except (KeyError, PermissionError) as e:
            logging.error(f"步骤 {step['id']} 失败: {e}")
            print(json.dumps({"status": "error", "failed_step": step["id"],
                              "error": str(e)}, ensure_ascii=False))
            sys.exit(1)
        if result.get("status") != "success":
            logging.error(f"步骤 {step['id']} 失败: {result.get('error')}")
            print(json.dumps({"status": "error", "failed_step": step["id"],
                              "error": result.get("error")}, ensure_ascii=False))
            sys.exit(1)
        context["steps"][step["id"]] = result
        print(f"[工作流] 第 {i}/{len(steps)} 步完成: {result.get('summary', '')}")

    # 汇总所有步骤的输出文件
    outputs = {}
    for res in context["steps"].values():
        outputs.update(res.get("outputs", {}))
    logging.info(f"===== 工作流完成: {wf.get('name')} =====")
    print(json.dumps({
        "status": "success",
        "workflow": wf.get("name"),
        "outputs": outputs,
        "summary": f"工作流完成，生成文件: {', '.join(outputs.values())}",
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()