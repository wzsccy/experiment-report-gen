---
name: experiment-report
description: "自动化深度学习实验流程：根据实验要求文件（.txt）生成PyTorch训练代码、运行训练、生成实验报告（.docx）。当用户提到实验报告、实验要求、深度学习实验、PyTorch实验、CNN/RNN/Transformer实验、图像分类/文本分类实验、训练代码时使用此技能。适用于需要从实验要求到最终报告全流程自动化的场景。"
---

# 深度学习实验自动化

将实验要求文件转换为可运行的 PyTorch 训练代码 + 标准格式的 docx 实验报告。

## 前置条件

需要全局安装 `experiment-report-gen` npm 包：

```bash
npm install -g experiment-report-gen
```

安装后可使用两个 CLI 命令：
- `experiment-report generate <config.json>` — 从配置文件生成 .docx 报告
- `experiment-report analyze <results.json> [--type image|text|auto]` — 从 results.json 自动生成分析文本

## 参考模板

参考已有的报告模板和封面模板，理解格式、内容深度和排版风格。

## 工作流程

```
实验要求.txt → 分析需求 → 生成/检查 train.py → 运行训练 → 收集报错日志 → 生成实验报告.docx
```

## Step 1: 分析实验要求

读取用户的实验要求文件（通常是 `实验X要求.txt`），提取：
- 实验标题和目的
- 数据集（CIFAR-10, AG_NEWS, MNIST 等）
- 模型类型（CNN, RNN, LSTM, Transformer 等）
- 任务类型（图像分类, 文本分类, 目标检测等）
- 具体输出要求（曲线图、预测结果可视化、准确率指标等）
- 思考题

## 环境要求

**所有 Python 命令必须在 conda 环境 `myvenv` 中运行，绝不能使用 base 环境。**

```bash
conda activate myvenv
```

后台或非交互模式下无法正确激活 conda 环境，需直接调用解释器完整路径：

```bash
/home/eleven/miniconda3/envs/myvenv/bin/python train.py 2>&1 | tee output/train.log
```

## Step 2: 生成训练代码

根据实验要求生成 `train.py`。代码必须包含：

### 必须输出的文件（保存到 `./output/`）
1. **`results.json`** — 包含所有训练指标的 JSON 文件
2. **`curves.png`** — Loss 和 Accuracy 曲线图
3. **任务特定可视化** — 如 `predictions.png` 或 `attention_weights.png`
4. **`best_model.pth`** — 最佳模型权重

### 运行训练
```bash
conda activate myvenv && python train.py 2>&1 | tee output/train.log
# 或后台运行（非交互模式，使用完整路径）：
/home/eleven/miniconda3/envs/myvenv/bin/python train.py 2>&1 | tee output/train.log
```
捕获所有输出，包括报错信息。如果目录已有 `output/results.json`，可以跳过训练直接生成报告。

## Step 3: 收集报错和运行日志

训练过程中的问题要在报告的"4.1 实验环境搭建与问题分析"中详细描述。常见问题：
- 环境问题：PyTorch 未安装、CUDA 不兼容、依赖缺失
- 数据问题：数据集路径错误、数据格式不匹配
- 运行问题：显存不足（OOM）、训练中断、后台任务失败
- 可视化问题：中文乱码、matplotlib 后端错误
- 性能问题：准确率不达标、过拟合、收敛慢

如果有 `output/train.log`，**必须读取其中的报错信息**，在报告中还原真实的报错过程和解决方法。

## Step 4: 生成实验报告

### 4a. 从 results.json 自动生成分析文本

```bash
experiment-report analyze <project-dir>/output/results.json --type image|text|auto
```

命令会自动检测实验类型（图像分类/文本分类）并输出分析文本到 stdout，可直接用于 config.json 的 sections.results。

### 4b. 准备 config.json（内容总字数不少于 5000 字）

**每个实验报告必须包含至少 3 个数据表格**，通过 config.json 的 `tables` 字段配置。常见表格类型：
- 模型结构对比表（不同模型的层配置、参数量）
- 超参数配置表（学习率、batch_size、优化器等）
- 数据集统计表（样本数、类别分布、划分比例）
- 性能对比表（准确率、损失、参数量等指标）
- 训练过程表（关键 epoch 的 loss/acc 数值）

表格数据必须从 `results.json` 中提取实际数值，不能编造。

**content（实验内容）必须包含以下子章节，总字数不少于 2000 字：**

```
4.1 实验环境搭建与问题分析（500+字）
    - 环境配置过程：Python版本、PyTorch版本、CUDA版本
    - 遇到的问题及报错：从 train.log 中提取实际报错信息，用引号标注具体错误消息
    - 解决方法：每个问题都详细描述排查过程和最终解决方案
    - 例如："ModuleNotFoundError: No module named 'torch'" → 通过 pip install 解决
    - 例如：后台运行训练静默终止 → 通过直接调用解释器路径解决
    - 例如：中文显示乱码 → 通过配置中文字体解决
    - CPU vs GPU 训练效率对比

4.2 数据集与预处理方法（300+字）
    - 数据集详细介绍：来源、规模、类别、图像/文本特点
    - 预处理步骤：归一化参数、数据增强策略（随机裁剪、翻转等）
    - 数据加载方式：DataLoader配置、batch_size选择

4.3 模型结构设计（400+字）
    - 网络架构详解：逐层描述每一层的类型、参数、作用
    - 关键组件说明：BatchNorm、Dropout、注意力机制等
    - 参数量统计
    - 为什么选择这个架构

4.4 训练策略与优化方法（400+字）
    - 超参数设置：学习率、batch_size、epochs、weight_decay
    - 优化器选择及理由
    - 学习率调度策略
    - 正则化方法：Dropout、L2正则化
    - 梯度裁剪等技巧

4.5 模型训练过程分析（400+字）
    - 训练初期：模型表现、损失变化
    - 训练中期：收敛速度、准确率提升
    - 训练后期：稳定状态、最终表现
    - 各阶段的特征描述
```

**results（实验结果）必须包含以下子章节，总字数不少于 1500 字：**

```
5.1 损失函数变化分析（300+字）
    - 引用 results.json 中的实际数值
    - 训练/测试损失的下降趋势（用具体数值描述）
    - 是否过拟合的判断依据

5.2 准确率变化分析（300+字）
    - 引用实际准确率数值
    - 准确率变化趋势描述
    - 最佳准确率及是否达标

5.3 分类/预测结果分析（300+字）
    - 描述预测图中的具体结果
    - 正确分类的分析
    - 错误分类的原因分析

5.4 模型泛化能力分析（200+字）
    - 训练集vs测试集差距分析
    - 正则化效果评估

5.5 模型优化策略效果分析（200+字）
    - BatchNorm的效果
    - 数据增强的效果
    - 学习率调度的效果

5.6 综合评价（100+字）
    - 整体性能总结
    - 是否达到实验要求
```

**thought_questions（思考题）不少于 500 字：**
- 逐题回答
- 结合实验数据和深度学习理论
- 有理有据，分点论述

### 4c. 调用 npm 包生成 docx

```bash
experiment-report generate <config-json-path>
```

### config.json 完整格式

```json
{
  "title": "实验标题",
  "student_name": "张三",
  "student_id": "20240000001",
  "class_name": "24数据科学与大数据技术",
  "teacher": "李四",
  "course": "课程名称",
  "date": "2026年5月3日",
  "semester": "2025-2026 学年第 2 学期",
  "output_path": "实验报告-实验四.docx",
  "sections": {
    "purpose": "400+字的实验目的...",
    "references": "参考文献...",
    "tasks": "实验任务...",
    "content": "2000+字，包含4.1-4.5子章节...",
    "results": "1500+字，包含5.1-5.6子章节...",
    "thought_questions": "500+字的思考题解答..."
  },
  "images": [
    { "path": "output/curves.png", "caption": "图 1 训练Loss与Accuracy曲线", "section": "results" },
    { "path": "output/predictions.png", "caption": "图 2 测试集预测结果", "section": "content" }
  ],
  "tables": [
    {
      "caption": "表 1 模型结构对比",
      "section": "content",
      "position": "start",
      "headers": ["组件", "模型A", "模型B"],
      "rows": [
        ["嵌入维度", "128", "64"],
        ["隐藏维度", "256", "128"]
      ]
    },
    {
      "caption": "表 2 性能对比",
      "section": "results",
      "headers": ["指标", "模型A", "模型B"],
      "rows": [
        ["准确率", "92.68%", "86.99%"]
      ]
    }
  ]
}
```

**图片配置说明：**
- `images` 是数组格式（不是对象）
- `section` 可选值：`content`（嵌入4.x节末尾）、`results`（嵌入5.x节末尾）
- 图片找不到时脚本会跳过并警告

**表格配置说明：**
- `tables` 是数组格式，每个表格包含 `caption`（标题）、`headers`（表头数组）、`rows`（数据二维数组）
- `section` 指定表格属于哪个章节：`content`（嵌入4.x节）、`results`（嵌入5.x节）
- `position` 可选：`start`（紧跟章节标题之后）、默认为 `end`（正文之后、图片之前）
- 表格自动带蓝色表头背景、居中对齐、边框
- **必须添加数据表格**：模型对比表、超参数表、数据集统计表、性能对比表、训练过程表等
- 从 `results.json` 中提取实际数值填充表格，确保数据准确

## 报告格式标准

### 封面页（纯段落，非表格）
参考 `实验报告样本.docx` 的段落格式：
- 学院名称（居中）
- 实 验 报 告（居中）
- 学期（居中）
- 课程/实验名称/班级/姓名学号/同组成员/评阅教师（左对齐，标签与值之间用空格分隔）
- 日期（居中）

### 正文页
- 实验标题（居中）
- 实验时间（左对齐）
- Table 0: 5行1列（一~五节），带边框
- Table 1: 1行1列（思考题），带边框
- 图片嵌入在表格单元格内，居中，下方附图注

## Step 5: 验证和交付

1. 确认报告总字数不少于 5000 字
2. 确认 docx 报告已生成
3. 告知用户报告路径

## 常见实验模式

### 图像分类（如 CIFAR-10）
- 模型：VGG-like CNN / ResNet
- 输出：curves.png + predictions.png

### 文本分类（如 AG_NEWS）
- 模型：LSTM / BiLSTM + Attention
- 输出：curves.png + attention_weights.png

### 对比实验
- 需要训练多个模型（如有/无某机制）
- 报告重点：对比分析、机制重要性论证
