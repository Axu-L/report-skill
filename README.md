# report-skill

实验报告生成技能（Claude Code Skill）：根据实验指导书、题目、代码或操作说明，自动填写信息工程学院 Word 实验报告模板，生成五个正文区域——实验目的、实验原理、实验内容及过程、实验结果及分析、实验小结，并按任务交付源码/原生工程文件。

## 目录结构

```
generate-lab-report/
├── SKILL.md                          # 技能主指令（排版、语言、工作流程）
├── agents/openai.yaml                # Agent 配置
├── assets/
│   ├── template.docx                 # 内置信息工程学院实验报告模板
│   └── icon.svg
├── references/
│   └── template-and-content.md       # 模板定位与 JSON 内容输入说明
└── scripts/
    ├── fill_report.py                # 按 JSON 填写模板生成报告
    └── check_layout.py               # 渲染后检查固定区域与栏目同页
```

## 使用方法

1. 把 `generate-lab-report` 放入技能目录，按 `SKILL.md` 工作流程读取实验指导书/题目/代码。
2. 按 `references/template-and-content.md` 的 JSON 格式准备五区内容：

```json
{
  "purpose": ["…"],
  "principle": ["…"],
  "process_pages": [["1. 题目原文", {"type":"code","text":"…"}, {"type":"screenshot","name":"任务1_运行结果.png","comment":"…"}]],
  "results": ["实验结果：…"],
  "summary": ["…"]
}
```

3. 生成报告：

```bash
python generate-lab-report/scripts/fill_report.py 内容.json 报告.docx
# 可选：--template 自定义模板.docx
python generate-lab-report/scripts/check_layout.py 模板.docx 报告.docx --pdf 报告.pdf
```

## 核心规则速览

- **题目原文不变**：任务标题与要求逐字沿用指导书，题号按出现顺序统一 renumber 为 1. 2. 3. 4. ……
- **实验视为已完成**：结果区统一写"实验结果"，用陈述句做结果分析（可填写合理实测数值），禁止"若…则…可能…"预测句式和"以…为准 / 依据记录填写 / 仅为示例 / 待填写"等提示化表述
- **去 AI 腔**：正文不出现"教师指定 / 按教师分配 / 老师要求"等要求腔，不出现"见附件 / 见附属文件"等附件腔
- **截图**：每个任务都必须有结果截图占位（`【截图：名称.png】`）+ 真实 Word 批注；"代码不要截图"仅指代码用 code 块文本呈现，运行结果仍需截图
- **排版**：正文宋体小四 12 磅、1.5 倍行距、首行缩进 2 字符；command 块 Consolas 12 磅左缩进 2 字符；code 块 Consolas 10 磅左缩进 2 字符、保留代码缩进
- **版式约束**：只填五个正文区域，固定栏位保留；首页含目的+原理，末页含结果分析+小结+态度+安全+批阅；过程写入唯一单元格跨页流动

## 交付目录约定

```
实验编号_实验名称/
├── 实验报告.docx
├── 任务01_名称/   # 完整源文件、配置、数据
├── 任务02_名称/
└── 交付说明.txt   # 软件版本、运行方法、检查结果、未验证项
```
