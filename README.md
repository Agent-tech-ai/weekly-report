# Agentech Weekly Report Template

这是 Wesley 每周工作汇报的可复用模板。  
你以后只需要更新 `config/weekly_report.json`，然后运行生成脚本，就能输出同一套格式的 PDF 和 Word 文件。

## 文件结构

```text
weekly-report/
  config/weekly_report.json      每周要填写的内容
  scripts/generate_report.py     生成 PDF / DOCX 的脚本
  outputs/                       生成后的周报文件
  examples/sample_polished_report.pdf
```

## 每周怎么用

1. 打开 `config/weekly_report.json`
2. 修改汇报周期：

```json
"period": {
  "start": "2026/06/15",
  "end": "2026/06/19"
}
```

3. 修改每个工作模块的进展：

```json
{
  "name": "摄像头标定与空间定位",
  "current_progress": "已完成摄像头安装高度确认，正在采集不同距离下的标定数据。",
  "completion": "60%",
  "notes": "下一步重点验证镜头畸变对测量结果的影响。"
}
```

4. 如果某个模块需要数量、次数、测试轮数等统计，也可以加可选字段：

```json
"current_count": "完成 6 轮移动测试"
```

5. 生成周报：

```bash
python3 scripts/generate_report.py
```

生成结果会出现在 `outputs/` 目录里：

```text
Agentech_周工作汇报_Wesley_2026-06-15_to_2026-06-19.pdf
Agentech_周工作汇报_Wesley_2026-06-15_to_2026-06-19.docx
```

## 以后你可以直接这样跟 Codex 说

```text
这周 Connie 网站上线 100%，摄像头标定 60%，人体识别 40%，机器人狗 Tilt 测试完成 6 轮，自主移动测试完成基础前进和转向，其他保持待填写。帮我生成本周周报。
```

Codex 会更新 `config/weekly_report.json` 并重新生成 PDF / DOCX。

## 安装依赖

如果本机缺依赖：

```bash
python3 -m pip install -r requirements.txt
```

macOS 一般可以直接生成中文 PDF。Windows/Linux 如果中文显示异常，安装 Noto Sans CJK 或微软雅黑字体后再运行。
