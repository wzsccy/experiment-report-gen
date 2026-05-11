---
name: experiment-report
description: "从 config.json 生成标准格式的 .docx 实验报告，支持实验报告模式和通用模式。当用户提到实验报告、docx报告、生成报告时使用此技能。"
---

# 实验报告生成器

从 JSON 配置文件生成标准格式的 `.docx` 报告。

## 前置条件

```bash
npm install -g experiment-report-gen
```

CLI 命令：
- `experiment-report generate <config.json>` — 从配置生成 .docx
- `experiment-report analyze <results.json> [--type image|text|auto]` — 自动生成分析文本

## 工作流程

```
config.json → experiment-report generate → .docx 报告
```

## Step 1: 确认报告模式

用 AskUserQuestion 询问报告模式：

| 选项 | 说明 | config 标志 |
|------|------|------------|
| 实验报告模式（默认） | 固定 4.1-4.5 + 5.1-5.6 结构，内置封面 | 无 `mode` 或 `"mode": "experiment"` |
| 通用模式 | 自由章节，适合非实验文档 | `"mode": "generic"` |
| 自定义模板 | 用自己的 .docx 合并封面 | `"cover_template": "<绝对路径>"` |

选"自定义模板"后追问路径。

## Step 2: 自动生成分析文本（实验报告模式）

```bash
experiment-report analyze <project-dir>/output/results.json --type image|text|auto
```

## Step 3: 准备 config.json（总字数 ≥ 5000）

**禁止使用 markdown 符号**：config.json 的所有文本字段中不得出现：`**`、`###`/`##`/`#`、`- `、`` ` ``、`>`

**至少 3 个数据表格**（模型对比/超参数/数据集统计/性能对比/训练过程），数据从 `results.json` 提取。

**content（≥ 2000 字）：**
- 4.1 环境搭建与问题分析（500+）：环境配置、实际报错（引号标注）、解决方法
- 4.2 数据集与预处理（300+）：数据来源/规模/类别、归一化/增强、DataLoader
- 4.3 模型结构设计（400+）：逐层描述、关键组件、参数量、架构选择理由
- 4.4 训练策略与优化（400+）：超参数、优化器、lr 调度、正则化、梯度裁剪
- 4.5 训练过程分析（400+）：初期/中期/后期的损失和指标变化

**results（≥ 1500 字）：**
- 5.1 损失函数分析（300+）：引用实际数值、下降趋势、过拟合判断
- 5.2 准确率/精度分析（300+）：引用数值、变化趋势、是否达标
- 5.3 分类/预测结果分析（300+）：预测图描述、正确/错误原因
- 5.4 泛化能力分析（200+）：训练 vs 测试差距、正则化效果
- 5.5 优化策略效果（200+）：BatchNorm/数据增强/lr 调度效果
- 5.6 综合评价（100+）：整体性能、是否达标

**thought_questions（≥ 500 字）：** 逐题回答，结合数据和理论，分点论述。

## Step 4: 生成 docx

```bash
experiment-report generate <config-json-path>
```

## config.json 完整格式（实验报告模式）

```json
{
  "mode": "experiment",
  "cover_template": "../experiment-report/templates/实验报告样本.docx",
  "title": "实验标题",
  "student_name": "张三",
  "student_id": "20240000001",
  "class_name": "24数据科学与大数据技术",
  "teacher": "李四",
  "course": "课程名称",
  "date": "2026年5月3日",
  "semester": "2025-2026 学年第 2 学期",
  "output_path": "实验报告-实验X.docx",
  "sections": {
    "purpose": "400+字...",
    "references": "参考文献...",
    "tasks": "实验任务...",
    "content": "2000+字，含4.1-4.5...",
    "results": "1500+字，含5.1-5.6...",
    "thought_questions": "500+字..."
  },
  "images": [
    { "path": "output/curves.png", "caption": "图 1 ...", "explain": "...", "section": "results" }
  ],
  "tables": [
    { "caption": "表 1 ...", "section": "content", "position": "start", "explain": "...", "headers": ["列1","列2"], "rows": [["值1","值2"]] }
  ]
}
```

## 通用模式 config 格式

```json
{
  "mode": "generic",
  "title": "文档标题",
  "date": "2026年5月5日",
  "output_path": "输出.docx",
  "cover": {
    "title": "课 程 报 告", "subtitle": "副标题",
    "fields": [{ "label": "学    院", "value": "计算机学院" }],
    "date": "2026年5月5日"
  },
  "body_sections": [
    {
      "title": "一、章节标题", "key": "key", "content": "正文...",
      "subsections": [{ "title": "1.1 子章节", "content": "..." }],
      "tables": [{ "caption": "表 1 ...", "headers": ["列1"], "rows": [["值1"]] }]
    }
  ],
  "images": [{ "path": "output/curves.png", "caption": "图 1 ...", "section": "results" }]
}
```

与实验报告模式区别：`body_sections` 数组代替 `sections`，封面用 `cover` 对象，无固定章节结构。
config 示例：`templates/config-example.json`、`templates/config-generic-example.json`

## 字段说明

**cover_template（仅实验报告模式）：**
- `null` → 自动生成封面；docx 路径 → 提取模板封面合并正文；通用模式不用此字段
- 路径相对于 config.json 所在目录，用 `/` 或 `\\`

**images：** 数组，`section` 可选 `content`/`results`，`explain` 可选，图片缺失时跳过。

**tables：** 数组，含 `caption`/`headers`/`rows`，`section` 指定嵌入章节，`position` 可选 `start`/`end`(默认)，`explain` 可选。数据必须从 `results.json` 提取。

## 报告格式标准

**封面页：** 模板封面 → 提取替换；自动生成 → 学院名(居中) + "实 验 报 告"(居中) + 学期(居中) + 课程/实验/班级/姓名学号/教师(左对齐) + 日期(居中)

**正文页：** 标题(居中) + 时间(左对齐) + Table 0(5行1列，一~五节) + Table 1(思考题) + 图片嵌入单元格

## 写作规则

生成报告时，整体风格应接近"学生完成实验后整理出来的正式报告"。优先使用自然连贯的大段描述，不要把正文写成大量机械的项目符号。每个小节应尽量形成完整段落，让内容看起来有实际实验过程，而不是由 AI 拼接出来的提纲。

语言要正式、清楚、具体，但不要过度学术化。技术内容应围绕本次实验本身展开，重点写实验中做了什么、为什么这样设置、运行后观察到了什么、结果有什么问题、如何分析这些现象。

优先采用"段落式叙述 + 必要表格 + 少量列表"的结构。不要把每一节都写成 1、2、3、4 的小点。不要每段都采用"首先、其次、再次、最后"的固定结构，也不要每段都以"本实验"开头。

结果分析必须围绕实际结果写，不能空泛总结。应优先引用用户提供的损失值、曲线变化等具体信息。如果实验结果一般，应如实写出不足。

**禁止 AI 化表达替换表：**

| 禁止写法 | 改为 |
|---------|------|
| 本实验旨在深入理解…… | 本次实验主要是练习…… |
| 具有重要意义 | 对理解相关内容有一定帮助 |
| 为后续学习奠定良好基础 | 对后续学习有一定参考 |
| 显著提升 | 有所改善 / 更稳定一些 |
| 充分证明 | 可以看出 / 从结果上看 |
| 整体质量优秀 | 整体效果尚可 / 多数结果可以辨认 |
| 完美实现 | 基本实现 |
| 动态平衡现象 | 两个损失逐渐趋于稳定 |
| 逼真的图像 | 较接近手写数字的图像 |
| 生成质量优秀 | 多数数字轮廓可以辨认 |
| 成功实现所有预定目标 | 实验内容基本完成 |

**去 AI 化自检清单（生成报告后必须逐项检查）：**

1. 检查是否存在过多模板化开头。如果多个段落都以"本实验""本文"开头，需要改写一部分。
2. 检查是否存在过度肯定的结论。如果出现"显著提升""充分证明"而没有足够数据支撑，应改成更保守的说法。
3. 检查是否重复解释同一概念。如果已在前文解释过，后文只需简单提到。
4. 检查段落是否过碎。如果某一节出现大量短小项目符号，应合并为自然段落。
5. 检查结果分析是否只写优点。实验结果部分必须包含一定的不足分析。
6. 检查内容是否像本科实验报告。最终文本应像学生完成实验后整理出的课程报告，而非论文或 AI 总结。

## Step 5: 验证

1. 确认报告 ≥ 5000 字
2. 确认 .docx 已生成
3. 告知用户报告路径
