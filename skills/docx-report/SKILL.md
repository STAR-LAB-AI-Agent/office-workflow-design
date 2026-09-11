\---

name: docx-report

description: 分析 Excel 数据并自动生成 Word 分析报告。当用户要求"分析数据并生成（Word）报告"时使用的两步工作流。

\---



\# 分析 Excel 并生成 Word 报告



\## 使用场景

用户希望得到一份正式的 Word 分析报告，而不只是口头结论。本技能自动执行两步工作流：

第1步 excel\_summary.py 生成统计摘要 → 第2步 json\_to\_docx.py 把摘要渲染成 Word（步骤间自动传递结果）。



\## 调用方式

cd C:\\Users\\Lenovo\\Desktop\\office-flow

python workflow.py workflows\\分析并生成Word报告.json --param excel=\&lt;用户指定的Excel路径\&gt;



\## 参数

\- --param excel=路径：要分析的 Excel。若用户未指定，默认使用 test\_data\\销售数据.xlsx。



\## 结果格式

引擎最后一行输出 JSON：{"status": "success", "outputs": {"docx": "报告路径"}, "summary": "..."}

\- 成功后告诉用户报告文件的完整路径和统计结论要点。

\- 若 status 为 error，failed\_step 字段表示失败的步骤，向用户说明哪一步出了问题。



\## 示例

1\. 用户："分析销售数据，帮我生成一份 Word 报告" → 执行上述命令，回复报告路径。

2\. 用户："把 test\_data\\销售数据.xlsx 做成分析报告" → 执行命令并带 --param excel=test\_data\\销售数据.xlsx。

