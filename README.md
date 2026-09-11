# office-flow：AI 办公任务流程编排

> 《智能体开发实战》课程实验 · 选题 27：AI 办公任务流程编排
> 将已有办公 Script/Skill 按用户目标组合成简单工作流，支持多个步骤之间的结果传递。

## 项目简介

用户用自然语言提出办公需求（如"分析销售数据并生成 Word 报告和 PPT"），nanobot 智能体
根据 SKILL.md 选择合适的能力，调用工作流引擎 workflow.py，按流程定义依次执行 Python
脚本，前一步的输出自动作为后一步的输入（结果传递）。

链路：用户自然语言 → nanobot → SKILL.md → workflow.py → Python Script → 开源库 → 结果

## 支持的三类核心意图

| # | 意图 | 示例说法 | 执行内容 |
|---|------|---------|---------|
| 1 | 数据分析 | "分析一下销售数据，哪个地区销售额最高？" | 单步：excel_summary.py |
| 2 | 分析+Word | "分析数据并生成一份 Word 报告" | 两步工作流 |
| 3 | 分析+Word+PPT | "生成 Word 报告和汇报 PPT" | 三步工作流 |

## 目录结构

```
office-flow/
├── scripts/                  # 可独立运行的办公脚本（统一 CLI 接口）
│   ├── excel_summary.py      #   Excel → 统计摘要 JSON（pandas）
│   ├── json_to_docx.py       #   摘要 JSON → Word 报告（python-docx）
│   ├── json_to_pptx.py       #   摘要 JSON → 汇报 PPT（python-pptx）
│   └── make_test_data.py     #   生成虚构测试数据
├── skills/                   # 供 nanobot 加载的 Skill（复制到 nanobot workspace）
├── workflows/                # 流程定义（JSON，声明步骤顺序与结果传递）
├── workflow.py               # 工作流引擎：顺序执行、占位符解析、日志、安全校验
├── tests/run_tests.py        # 10 个自动化测试用例（正常/边界/异常）
├── test_data/                # 虚构测试数据（无真实个人信息）
├── outputs/                  # 生成的结果文件
├── logs/                     # 运行日志（不含敏感信息）
└── requirements.txt
```

## 统一 Script 接口约定

所有脚本遵循同一接口：
- 入参：`--input &lt;文件&gt; --output &lt;文件&gt;`
- 出参：向标准输出打印一行 JSON：
  `{"status": "success|error", "outputs": {...}, "summary": "一句话结果"}`
- 失败时 `status=error` 并给出可读原因，exit code 非 0

## 工作流定义与结果传递

`workflows/*.json` 声明步骤；占位符 `{{steps.&lt;步骤id&gt;.outputs.&lt;字段&gt;}}` 在执行时
被替换为前面步骤的真实输出：

```json
{"id": "docx", "script": "scripts/json_to_docx.py",
 "args": {"input": "{{steps.summary.outputs.summary_json}}", "output": "outputs/分析报告.docx"}}
```

`--param key=value` 可在运行时覆盖流程参数（如指定不同的 Excel 文件）。

## 安装与运行

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install nanobot-ai        # 智能体运行环境（也可用其他 Agent Runtime）
```

脚本独立运行（无需智能体，便于单独测试）：

```powershell
python scripts\excel_summary.py --input test_data\销售数据.xlsx --output outputs\摘要.json
python workflow.py workflows\分析并生成Word报告和PPT.json --param excel=test_data\销售数据.xlsx
```

nanobot 集成：将 `skills/` 下各目录复制到 `~/.nanobot/workspace/skills/`，
然后 `nanobot agent -m "分析销售数据并生成 Word 报告和 PPT"`。

## 测试

```powershell
python tests\run_tests.py
```

10 个用例：正常 5 个（单脚本×3、完整工作流、数据生成）、异常 3 个（文件不存在、
参数格式错误、路径越界）、边界 2 个（空 Excel、纯文本 Excel）。当前全部 PASS。

## 开源依赖与许可证

| 项目 | 版本 | 许可证 | 用途 |
|------|------|--------|------|
| pandas | ≥2.0 | BSD-3-Clause | Excel 读取与统计汇总 |
| openpyxl | ≥3.1 | MIT | Excel 文件格式支持（pandas 引擎） |
| python-docx | ≥1.1 | MIT | Word 报告生成 |
| python-pptx | ≥1.0 | MIT | PPT 生成 |
| HKUDS/nanobot | latest | MIT | 智能体入口与 Skill 编排层 |

## 安全设计

- 最小权限：工作流引擎校验所有文件路径必须位于项目目录内，越界即拒绝；
- 敏感信息：DeepSeek API Key 仅保存在本机 `~/.nanobot/config.json`，不进入项目与 Git；
- 日志：logs/workflow.log 记录每次执行的时间、命令与结果，不记录任何凭据；
- 失败即停：任一步骤失败立即终止流程并指明 failed_step，不产生半成品输出。

## 低 Token 优化

原始 Excel（120 行 × 6 列）不进入模型上下文：由 pandas 在本地完成统计，
仅把约 200 字的结构化摘要 JSON 返回给智能体，相比直接读文件 Token 消耗降低约 90%。

## 已知问题

- nanobot 退出时 HTTP 依赖库偶发打印无害的清理警告（不影响任务结果）；
- 仅支持 .xlsx 格式 Excel。
