\---

name: office-workflow

description: 分析 Excel 并同时生成 Word 报告和汇报 PPT。当用户要求"生成报告和 PPT / 全套汇报材料"时使用的三步工作流编排。

\---



\# 办公任务流程编排：分析 → Word 报告 → PPT



\## 使用场景

用户要求一次性产出汇报材料（Word 报告 + PPT）。本技能按流程定义依次执行：

summary（分析Excel）→ docx（生成Word）→ pptx（生成PPT），后两步共用第1步的摘要结果。



\## 调用方式

cd C:\\Users\\Lenovo\\Desktop\\office-flow

python workflow.py workflows\\分析并生成Word报告和PPT.json --param excel=\&lt;用户指定的Excel路径\&gt;



\## 参数

\- --param excel=路径：要分析的 Excel。若用户未指定，默认使用 test\_data\\销售数据.xlsx。



\## 结果格式

引擎最后一行输出 JSON：{"status": "success", "outputs": {"summary\_json": "...", "docx": "...", "pptx": "..."}}

\- 成功后把生成的文件列表和核心统计结论告诉用户。

\- 若某步失败（status=error），failed\_step 指明失败步骤，向用户说明并停止，不要重复执行。



\## 示例

1\. 用户："分析销售数据，生成 Word 报告和汇报 PPT" → 执行命令，回复两个文件路径。

2\. 用户："帮我准备这份数据的汇报材料" → 执行命令（即 Word + PPT 全套），总结关键结论。

