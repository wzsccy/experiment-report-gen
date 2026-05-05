# experiment-report-gen

通用报告生成器 — 从 JSON 配置文件生成标准格式的 `.docx` 报告。支持两种模式：

- **通用模式** (`mode: "generic"`) — 完全自定义封面、章节结构，适用于任何类型的报告
- **实验报告模式** (默认) — 预设的实验报告格式，向后兼容

## 安装

```bash
npm install -g experiment-report-gen
```

## 功能

- 从 `config.json` 生成格式化的 Word 报告（含封面、正文、表格、图片）
- 从 `results.json` 自动分析训练结果，生成报告文本
- 支持图像分类和文本分类两种实验类型的自动分析
- 支持 CLI 命令行和 Node.js 编程两种使用方式
- 支持自定义封面、章节、子章节、表格、图片

## CLI 用法

### 生成报告

```bash
experiment-report generate config.json
```

### 分析训练结果

```bash
# 自动检测实验类型
experiment-report analyze output/results.json

# 指定类型
experiment-report analyze output/results.json --type image
experiment-report analyze output/results.json --type text
```

## 编程用法

```javascript
const { generateReport } = require('experiment-report-gen');
generateReport('config.json').then(() => console.log('done'));

const { analyzeResults } = require('experiment-report-gen/src/analyze-results');
const text = analyzeResults('output/results.json', 'image');
```

---

## 通用模式 (mode: "generic")

适用于任何类型的报告：课程报告、调研报告、项目总结等。封面和章节结构完全可配置。

```json
{
  "mode": "generic",
  "title": "报告标题",
  "date": "2026年5月5日",
  "output_path": "报告.docx",
  "cover": {
    "title": "课 程 报 告",
    "subtitle": "副标题",
    "fields": [
      { "label": "姓    名", "value": "张三" },
      { "label": "学    号", "value": "20240000001" }
    ],
    "date": "2026年5月5日"
  },
  "body_sections": [
    {
      "title": "一、概述",
      "key": "overview",
      "content": "章节正文内容..."
    },
    {
      "title": "二、分析",
      "key": "analysis",
      "content": "正文内容...",
      "subsections": [
        {
          "title": "2.1 子章节标题",
          "content": "子章节内容..."
        }
      ],
      "tables": [
        {
          "caption": "表 1 数据对比",
          "headers": ["指标", "值"],
          "rows": [["准确率", "92.5%"]]
        }
      ]
    }
  ],
  "images": [
    { "path": "output/chart.png", "caption": "图 1 趋势图", "section": "analysis" }
  ]
}
```

完整示例见 `templates/config-generic-example.json`。

### 封面配置 (cover)

| 字段 | 说明 |
|------|------|
| `title` | 封面大标题，如"实验报告"、"课程报告" |
| `subtitle` | 副标题 |
| `titleSize` | 标题字号，默认 44 |
| `fields` | 信息字段数组，每项为 `{label, value}` 或 `{text, size, bold, align}` |
| `date` | 底部日期 |
| `spacer` | 顶部留白行数，默认 3 |

### 章节配置 (body_sections)

每个章节对象：

| 字段 | 说明 |
|------|------|
| `title` | 章节标题 |
| `key` | 唯一标识，用于关联图片和表格 |
| `content` | 正文内容，`\n\n` 分段 |
| `subsections` | 子章节数组，每项含 `title`、`content` |
| `tables` | 本章节的表格数组 |

---

## 实验报告模式 (默认)

预设的实验报告格式，包含固定的五章节结构和思考题。

```json
{
  "title": "实验标题",
  "student_name": "姓名",
  "student_id": "学号",
  "class_name": "班级",
  "teacher": "评阅教师",
  "course": "课程名称",
  "college": "学院名称",
  "date": "2026年5月5日",
  "semester": "2025-2026 学年第 2 学期",
  "output_path": "实验报告.docx",
  "sections": {
    "purpose": "实验目的和要求",
    "references": "相关资料和参考文献",
    "tasks": "实验任务",
    "content": "实验内容（步骤）",
    "results": "实验结果及分析",
    "thought_questions": "相关思考题及解答"
  },
  "images": [
    { "path": "output/curves.png", "caption": "图 1 训练曲线", "section": "results" }
  ],
  "tables": [
    {
      "caption": "表 1 数据对比",
      "section": "results",
      "headers": ["指标", "值"],
      "rows": [["准确率", "92.5%"]]
    }
  ]
}
```

完整示例见 `templates/config-example.json`。

### 字段说明

| 字段 | 必填 | 说明 |
|------|------|------|
| `title` | 是 | 实验标题 |
| `student_name` | 是 | 学生姓名 |
| `student_id` | 是 | 学号 |
| `class_name` | 是 | 班级 |
| `teacher` | 是 | 评阅教师 |
| `course` | 是 | 课程名称 |
| `college` | 否 | 学院名称，默认"学院" |
| `date` | 是 | 实验日期 |
| `semester` | 是 | 学期 |
| `output_path` | 否 | 输出路径，默认"实验报告.docx" |
| `sections` | 是 | 各章节内容 |
| `images` | 否 | 图片配置数组 |
| `tables` | 否 | 表格配置数组 |

---

## 通用配置

### 图片配置

- `path`: 图片路径（相对于 config.json 所在目录）
- `caption`: 图注文字
- `section`: 所属章节的 key
- `width` / `height`: 图片尺寸（像素），默认 450x300

### 表格配置

- `caption`: 表格标题
- `headers`: 表头数组
- `rows`: 数据二维数组
- `section`: 所属章节的 key（通用模式下也可写在章节对象内）
- `position`: 位置，`start`（章节标题后）或默认 `end`（正文后、图片前）

### 页面配置

- `font`: 全局字体，默认"宋体"
- `page_width` / `page_height`: 页面尺寸（DXA），默认 A4 (11906x16838)
- `margin`: 页边距（DXA），默认 1440
- `header`: 页眉文字，默认使用 title
- `output_path`: 输出文件路径

## 依赖

- [docx](https://www.npmjs.com/package/docx) - Word 文档生成库

## 许可证

MIT
