\---

name: excel-summary

description: 分析 Excel 文件并生成统计摘要。当用户想了解某个 Excel 的数据概况、统计指标、排名（如"哪个地区销售额最高"）时使用。

\---



\# Excel 数据统计摘要



\## 使用场景

用户想快速了解 Excel 数据概况（行数、合计、平均值、按类别汇总排名）时使用本技能。



\## 调用方式

先切换到项目目录，再执行脚本（两步）：

cd C:\\Users\\Lenovo\\Desktop\\office-flow

python scripts\\excel\_summary.py --input \&lt;用户指定的Excel路径\&gt; --output outputs\\摘要.json



\## 参数

\- --input：Excel 文件路径，必填。若用户未指定，先追问用户要分析哪个文件，不要擅自假设。

\- --output：固定使用 outputs\\摘要.json



\## 结果格式

脚本最后一行输出 JSON：{"status": "success", "outputs": {"summary\_json": "..."}, "summary": "一句话统计结论"}

\- 将 summary 的内容用自然语言转告用户。

\- 若 status 为 error，把 error 内容友好转告用户（如"文件不存在"），不要编造数据。



\## 示例

1\. 用户："帮我分析一下 test\_data\\销售数据.xlsx" → 执行命令，转告统计结论。

2\. 用户："哪个地区的销售额最高？" → 执行命令，从 summary 中提取"按地区汇总最高的是 XX"回答。

