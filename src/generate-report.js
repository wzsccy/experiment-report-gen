/**
 * generate-report.js - 通用报告生成器 (.docx)
 *
 * 支持两种模式:
 *   1. 通用模式 (mode: "generic") - 完全自定义封面、章节结构
 *   2. 实验报告模式 (默认) - 向后兼容原有实验报告格式
 *
 * v2.0 新增:
 *   - cover_template: 使用模板 docx 的封面页
 *   - explain: 表格和图片后的解释文字
 *   - 表格无蓝色表头，居中显示
 *
 * 用法:
 *   const { generateReport } = require('experiment-report-gen');
 *   generateReport('config.json').then(() => console.log('done'));
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, BorderStyle, WidthType,
  ShadingType, PageBreak, PageNumber, Header, Footer, TabStopType, TabStopPosition
} = require('docx');

// ── 辅助函数 ──────────────────────────────────────────
const SINGLE_BORDER = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const SINGLE_BORDERS = { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER };
const NONE_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NONE_BORDERS = { top: NONE_BORDER, bottom: NONE_BORDER, left: NONE_BORDER, right: NONE_BORDER };

const FONT = { ascii: 'Times New Roman', eastAsia: '宋体' };

// ── 外层模块表格常量 ─────────────────────────────────
const MODULE_TABLE_WIDTH = 9000;
const MODULE_CELL_MARGINS = { top: 80, bottom: 80, left: 150, right: 150 };
// 内容数据表格宽度 = 模块宽度 - 单元格左右边距
const DATA_TABLE_WIDTH = 8700;

/**
 * 创建外层模块行（章节容器行）
 * 注意：不设置 cantSplit，允许内容自然跨页
 */
function createOuterSectionRow(cellChildren) {
  return new TableRow({
    children: [
      new TableCell({
        borders: SINGLE_BORDERS,
        width: { size: MODULE_TABLE_WIDTH, type: WidthType.DXA },
        verticalAlign: 'top',
        margins: MODULE_CELL_MARGINS,
        children: cellChildren,
      }),
    ],
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.spaceBefore || 0, after: opts.spaceAfter || 0, line: opts.lineSpacing || 360 },
    children: [new TextRun({ text, font: opts.font || FONT, size: opts.size || 24, bold: opts.bold || false })],
  });
}

function pMulti(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.spaceBefore || 0, after: opts.spaceAfter || 0, line: opts.lineSpacing || 360 },
    children: runs.map(r => {
      if (typeof r === 'string') return new TextRun({ text: r, font: FONT, size: 24 });
      return new TextRun({ font: FONT, size: 24, ...r });
    }),
  });
}

function emptyLine() { return p(''); }

function coverInfoLine(label, value) {
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 360 },
    children: [
      new TextRun({ text: label, font: FONT, size: 24 }),
      new TextRun({ text: '\t' + (value || ''), font: FONT, size: 24 }),
    ],
    tabStops: [{ type: TabStopType.LEFT, position: 2400 }],
  });
}

/**
 * 清理文本中的 markdown 符号，避免它们出现在 docx 中。
 * - `### ` / `## ` / `# ` 行首标题标记
 * - `**text**` 加粗标记 → 保留 text
 * - `- ` 行首列表标记
 * - 行首 `1. ` / `2. ` 等编号列表标记（仅匹配 1 位数字 + 点 + 空格，避免误删 "4.1"）
 *
 * 注意：不用 regex 匹配加粗，因为 Node.js 的 . 在某些版本
 * 无法正确匹配 CJK 字符，改用 split 方案。
 */
function stripMarkdown(text) {
  if (!text) return '';
  let result = text
    // 行首标题标记
    .replace(/^#{1,4}\s+/gm, '')
    // 行首无序列表标记 "- "
    .replace(/^-\s+/gm, '');
  // 加粗标记 **text** → text（split 方案，避免 Node.js regex CJK bug）
  if (result.includes('**')) {
    result = result.split('**').join('');
  }
  // 行首有序列表标记 "1. " ~ "9. "（不匹配 "4.1" 这种小节号）
  result = result.replace(/^(\d)\.\s+/gm, '$1. ');
  return result;
}

function textToParagraphs(text, fontSize = 24) {
  const result = [];
  const cleanText = stripMarkdown(text);
  const blocks = cleanText.split('\n\n');
  for (const block of blocks) {
    const lines = block.split('\n');
    for (const line of lines) {
      if (line.trim()) {
        result.push(new Paragraph({
          spacing: { after: 80, line: 360 },
          indent: { firstLine: 480 },
          children: [new TextRun({ text: line.trim(), font: FONT, size: fontSize })],
        }));
      }
    }
  }
  return result;
}

/**
 * 将内容文本按 "4.x" / "5.x" 小节标题拆分，标题加粗，正文缩进。
 */
function parseContentSubsections(text, fontSize = 24) {
  if (!text) return [];
  const cleanText = stripMarkdown(text);
  const result = [];
  const parts = cleanText.split(/\n(?=\d+\.\d+\s+\S)/);

  for (const part of parts) {
    const trimmed = part.trim();
    if (!trimmed) continue;

    const lines = trimmed.split('\n');
    const firstLine = lines[0].trim();
    const isHeader = /^\d+\.\d+\s+/.test(firstLine);

    if (isHeader) {
      result.push(new Paragraph({
        spacing: { before: 160, after: 80, line: 360 },
        children: [new TextRun({ text: firstLine, bold: true, font: FONT, size: fontSize })],
      }));
      for (let i = 1; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line) {
          result.push(new Paragraph({
            spacing: { after: 80, line: 360 },
            indent: { firstLine: 480 },
            children: [new TextRun({ text: line, font: FONT, size: fontSize })],
          }));
        }
      }
    } else {
      for (const line of lines) {
        if (line.trim()) {
          result.push(new Paragraph({
            spacing: { after: 80, line: 360 },
            indent: { firstLine: 480 },
            children: [new TextRun({ text: line.trim(), font: FONT, size: fontSize })],
          }));
        }
      }
    }
  }
  return result;
}

function explainParagraph(text) {
  if (!text) return [];
  return [new Paragraph({
    spacing: { after: 200, line: 360 },
    indent: { firstLine: 480 },
    children: [new TextRun({ text, font: FONT, size: 24 })],
  })];
}

// ── 内容数据表格（正文中的表1、表2等，非外层模块框）──────────

function createDataTable(tableConfig) {
  const { caption, headers, rows, explain } = tableConfig;
  const paragraphs = [];

  if (caption) {
    paragraphs.push(new Paragraph({
      spacing: { before: 120, after: 60 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: caption, bold: true, font: FONT, size: 21 })],
    }));
  }

  const colCount = headers.length;
  const colWidth = Math.floor(DATA_TABLE_WIDTH / colCount);
  const colWidths = Array(colCount).fill(colWidth);
  colWidths[colCount - 1] = DATA_TABLE_WIDTH - colWidth * (colCount - 1);

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, idx) => new TableCell({
      borders: SINGLE_BORDERS,
      width: { size: colWidths[idx], type: WidthType.DXA },
      verticalAlign: 'center',
      margins: { top: 40, bottom: 40, left: 60, right: 60 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0, line: 276 },
        children: [new TextRun({ text: h, bold: true, font: FONT, size: 21 })],
      })],
    })),
  });

  const dataRows = rows.map(row =>
    new TableRow({
      children: row.map((cell, idx) => new TableCell({
        borders: SINGLE_BORDERS,
        width: { size: colWidths[idx], type: WidthType.DXA },
        verticalAlign: 'center',
        margins: { top: 40, bottom: 40, left: 60, right: 60 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 0, line: 276 },
          children: [new TextRun({ text: String(cell), font: FONT, size: 21 })],
        })],
      })),
    })
  );

  const table = new Table({
    width: { size: DATA_TABLE_WIDTH, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
    borders: SINGLE_BORDERS,
  });

  paragraphs.push(table);

  // 解释文字
  paragraphs.push(...explainParagraph(explain));

  return paragraphs;
}

// ── 图片加载 ──────────────────────────────────────────

function loadImage(imgConfig, configDir) {
  if (!imgConfig) return [];
  const imgPath = imgConfig.path;
  if (!imgPath) return [];

  const fullPath = path.resolve(configDir, imgPath);
  if (!fs.existsSync(fullPath)) {
    console.warn(`警告: 图片文件不存在: ${fullPath}`);
    return [];
  }

  const ext = path.extname(imgPath).toLowerCase().replace('.', '');
  const typeMap = { png: 'png', jpg: 'jpeg', jpeg: 'jpeg', gif: 'gif', bmp: 'bmp' };
  const imgType = typeMap[ext] || 'png';

  const data = fs.readFileSync(fullPath);
  const caption = imgConfig.caption || '';
  const width = imgConfig.width || 450;
  const height = imgConfig.height || 300;

  const paragraphs = [];
  // 图片段落：直接用段前间距，不用空段落
  paragraphs.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 60 },
    children: [
      new ImageRun({
        type: imgType,
        data,
        transformation: { width, height },
        altText: { title: caption, description: caption, name: caption },
      }),
    ],
  }));

  if (caption) {
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: caption, font: FONT, size: 20 })],
    }));
  }

  // 解释文字
  paragraphs.push(...explainParagraph(imgConfig.explain));

  return paragraphs;
}

// ── 通用模式: 构建封面 ──────────────────────────────
function buildGenericCover(config) {
  const cover = config.cover || {};
  const children = [];

  for (let i = 0; i < (cover.spacer || 3); i++) children.push(emptyLine());

  if (cover.title) {
    children.push(p(cover.title, {
      size: cover.titleSize || 44, bold: true, align: AlignmentType.CENTER, spaceAfter: 400
    }));
  }

  if (cover.subtitle) {
    children.push(p(cover.subtitle, {
      size: cover.subtitleSize || 28, align: AlignmentType.CENTER, spaceAfter: 300
    }));
  }

  if (Array.isArray(cover.fields)) {
    for (const field of cover.fields) {
      if (typeof field === 'string') {
        children.push(p(field, { size: 24, align: AlignmentType.CENTER, spaceAfter: 100 }));
      } else if (field.label) {
        children.push(coverInfoLine(field.label, field.value));
      } else if (field.text) {
        children.push(p(field.text, {
          size: field.size || 24,
          bold: field.bold || false,
          align: field.align || AlignmentType.CENTER,
          spaceAfter: field.spaceAfter || 100,
        }));
      }
    }
  }

  if (cover.date) {
    children.push(emptyLine());
    children.push(emptyLine());
    children.push(p(cover.date, { size: 24, align: AlignmentType.CENTER }));
  }

  return children;
}

// ── 实验报告模式: 构建封面 (向后兼容) ──────────────
function buildExperimentCover(config) {
  const { title, student_name, student_id, class_name, teacher, course, college, date, semester } = config;
  const children = [];

  children.push(emptyLine());
  children.push(emptyLine());
  children.push(emptyLine());
  children.push(p(college || '学院', {
    size: 36, bold: true, align: AlignmentType.CENTER, spaceAfter: 200
  }));
  children.push(p('实 验 报 告', {
    size: 44, bold: true, align: AlignmentType.CENTER, spaceAfter: 400
  }));
  children.push(p(semester || '', {
    size: 24, align: AlignmentType.CENTER, spaceAfter: 600
  }));
  children.push(coverInfoLine('课程（模块）', course));
  children.push(coverInfoLine('实验名称', title));
  children.push(coverInfoLine('班    级', class_name));
  children.push(coverInfoLine('姓名(学号)', `${student_name}（${student_id}）`));
  children.push(coverInfoLine('同组成员', ''));
  children.push(coverInfoLine('评阅教师', teacher));
  children.push(emptyLine());
  children.push(emptyLine());
  children.push(p(date || '', { size: 24, align: AlignmentType.CENTER }));

  return children;
}

// ── 通用模式: 构建正文 ──────────────────────────────
function buildGenericBody(config, sectionImages, sectionTables) {
  const bodySections = config.body_sections || [];
  const children = [];

  if (config.title) {
    children.push(p(config.title, {
      size: 32, bold: true, align: AlignmentType.CENTER, spaceBefore: 0, spaceAfter: 100
    }));
  }

  if (config.date) {
    children.push(p(`日期：${config.date}`, { size: 24, spaceAfter: 60 }));
  }

  const sectionRows = bodySections.map((section, i) => {
    const key = section.key || `section_${i}`;
    const cellChildren = [];

    const sectionTitle = section.title || `第${i + 1}节`;
    cellChildren.push(new Paragraph({
      spacing: { after: 80, line: 360 },
      children: [new TextRun({ text: sectionTitle, bold: true, font: FONT, size: 24 })],
    }));

    const tablesForSection = sectionTables[key] || (section.tables || []);
    for (const tbl of tablesForSection) {
      if (tbl.position === 'start') {
        cellChildren.push(...createDataTable(tbl));
      }
    }

    if (section.content) {
      cellChildren.push(...textToParagraphs(section.content));
    }

    for (const tbl of tablesForSection) {
      if (tbl.position !== 'start') {
        cellChildren.push(...createDataTable(tbl));
      }
    }

    const imgs = sectionImages[key] || [];
    if (imgs.length > 0) {
      cellChildren.push(...imgs);
    }

    if (Array.isArray(section.subsections)) {
      for (const sub of section.subsections) {
        cellChildren.push(new Paragraph({
          spacing: { before: 160, after: 80, line: 360 },
          children: [new TextRun({ text: sub.title, bold: true, font: FONT, size: 24 })],
        }));
        if (sub.content) {
          cellChildren.push(...textToParagraphs(sub.content));
        }
        if (Array.isArray(sub.tables)) {
          for (const tbl of sub.tables) {
            cellChildren.push(...createDataTable(tbl));
          }
        }
        if (Array.isArray(sub.images)) {
          for (const img of sub.images) {
            cellChildren.push(...loadImage(img, config._configDir));
          }
        }
      }
    }

    return createOuterSectionRow(cellChildren);
  });

  if (sectionRows.length > 0) {
    const mainTable = new Table({
      width: { size: MODULE_TABLE_WIDTH, type: WidthType.DXA },
      alignment: AlignmentType.CENTER,
      columnWidths: [MODULE_TABLE_WIDTH],
      rows: sectionRows,
      borders: SINGLE_BORDERS,
    });
    children.push(mainTable);
  }

  return children;
}

// ── 实验报告模式: 构建正文 (向后兼容) ──────────────

function buildExperimentBody(config, sectionImages, sectionTables) {
  const { title, date, sections } = config;
  const children = [];

  // 标题段落
  children.push(p(title, {
    size: 32, bold: true, align: AlignmentType.CENTER, spaceBefore: 0, spaceAfter: 100
  }));
  children.push(p(`实验时间：${date || ''}`, { size: 24, spaceAfter: 60 }));

  const sectionTitles = [
    '一、实验目的和要求',
    '二、相关资料和参考文献',
    '三、实验任务',
    '四、实验内容（步骤）',
    '五、实验结果及分析',
  ];
  const sectionKeys = ['purpose', 'references', 'tasks', 'content', 'results'];

  const sectionRows = sectionTitles.map((titleText, i) => {
    const key = sectionKeys[i];
    const content = sections[key] || '';
    const cellChildren = [];

    // 章节标题段落
    cellChildren.push(new Paragraph({
      spacing: { after: 80, line: 360 },
      children: [new TextRun({ text: titleText, bold: true, font: FONT, size: 24 })],
    }));

    // position=start 的内容表格
    const tablesForSection = sectionTables[key] || [];
    for (const tbl of tablesForSection) {
      if (tbl.position === 'start') {
        cellChildren.push(...createDataTable(tbl));
      }
    }

    // 正文内容
    if (key === 'content' || key === 'results') {
      cellChildren.push(...parseContentSubsections(content));
    } else {
      cellChildren.push(...textToParagraphs(content));
    }

    // position!=start 的内容表格
    for (const tbl of tablesForSection) {
      if (tbl.position !== 'start') {
        cellChildren.push(...createDataTable(tbl));
      }
    }

    // 图片
    if (sectionImages[key] && sectionImages[key].length > 0) {
      cellChildren.push(...sectionImages[key]);
    }

    return createOuterSectionRow(cellChildren);
  });

  // 思考题：作为 mainTable 的最后一行，不是单独的表格
  const thoughtCellChildren = [];
  thoughtCellChildren.push(new Paragraph({
    spacing: { after: 80, line: 360 },
    children: [new TextRun({ text: '相关思考题及解答', bold: true, font: FONT, size: 24 })],
  }));
  if (sections.thought_questions) {
    thoughtCellChildren.push(...textToParagraphs(sections.thought_questions));
  }
  sectionRows.push(createOuterSectionRow(thoughtCellChildren));

  // 一个主表格包含所有外层模块（含思考题）
  const mainTable = new Table({
    width: { size: MODULE_TABLE_WIDTH, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    columnWidths: [MODULE_TABLE_WIDTH],
    rows: sectionRows,
    borders: SINGLE_BORDERS,
  });
  children.push(mainTable);

  return children;
}

// ── 封面模板合并（调用 Python 脚本）──────────────────

function mergeCoverTemplate(templatePath, bodyDocxPath, outputPath) {
  const scriptDir = path.dirname(path.resolve(__filename || __dirname));
  const scriptPath = path.join(scriptDir, 'merge-cover.py');

  if (!fs.existsSync(scriptPath)) {
    console.warn(`警告: merge-cover.py 未找到，跳过封面模板合并`);
    return false;
  }

  try {
    execSync(`python3 "${scriptPath}" "${templatePath}" "${bodyDocxPath}" "${outputPath}"`, {
      stdio: 'pipe',
      timeout: 30000,
    });
    return true;
  } catch (err) {
    try {
      execSync(`python "${scriptPath}" "${templatePath}" "${bodyDocxPath}" "${outputPath}"`, {
        stdio: 'pipe',
        timeout: 30000,
      });
      return true;
    } catch (err2) {
      console.error(`封面模板合并失败: ${err2.message}`);
      return false;
    }
  }
}

// ── 主函数 ──────────────────────────────────────────

async function generateReport(configPath) {
  const configDir = path.dirname(path.resolve(configPath));
  const config = JSON.parse(fs.readFileSync(configPath, 'utf-8'));
  config._configDir = configDir;

  const isGeneric = config.mode === 'generic';
  const font = config.font || { ascii: 'Times New Roman', eastAsia: '宋体' };
  const coverTemplate = config.cover_template || null;

  // ── 收集图片 ──────────────────────────────────────
  const sectionImages = {};
  const images = config.images || [];
  for (const imgConfig of images) {
    const section = imgConfig.section || 'results';
    const imgParagraphs = loadImage(imgConfig, configDir);
    if (imgParagraphs.length > 0) {
      if (!sectionImages[section]) sectionImages[section] = [];
      sectionImages[section].push(...imgParagraphs);
    }
  }

  // ── 收集表格 ──────────────────────────────────────
  const sectionTables = {};
  const tables = config.tables || [];
  for (const tbl of tables) {
    const section = tbl.section || 'results';
    if (!sectionTables[section]) sectionTables[section] = [];
    sectionTables[section].push(tbl);
  }

  // ── 构建封面 ──────────────────────────────────────
  const coverChildren = isGeneric
    ? buildGenericCover(config)
    : buildExperimentCover(config);

  // ── 构建正文 ──────────────────────────────────────
  let bodyChildren;
  if (isGeneric) {
    bodyChildren = buildGenericBody(config, sectionImages, sectionTables);
  } else {
    bodyChildren = buildExperimentBody(config, sectionImages, sectionTables);
  }

  // ── 页眉页脚 ──────────────────────────────────────
  const headerText = config.header || config.title || '';

  // ── 组装文档 ──────────────────────────────────────
  const pageWidth = config.page_width || 11906;
  const pageHeight = config.page_height || 16838;
  const baseMargin = config.margin || 1440;
  // 正文区域页边距：top 加大，避免跨页后内容贴住页眉
  const topMargin = config.top_margin || 1800;
  const bottomMargin = config.bottom_margin || baseMargin;
  const leftMargin = config.left_margin || baseMargin;
  const rightMargin = config.right_margin || baseMargin;
  // 页眉/页脚距页面边缘距离
  const headerMargin = config.header_margin || 720;
  const footerMargin = config.footer_margin || 720;

  const pageMargin = {
    top: topMargin,
    right: rightMargin,
    bottom: bottomMargin,
    left: leftMargin,
    header: headerMargin,
    footer: footerMargin,
  };

  const bodySection = {
    properties: {
      page: {
        size: { width: pageWidth, height: pageHeight },
        margin: pageMargin,
      },
    },
    headers: {
      default: new Header({
        children: [p(headerText, { size: 18, align: AlignmentType.CENTER })],
      }),
    },
    footers: {
      default: new Footer({
        children: [pMulti(
          [{ text: '第 ', size: 18 }, { children: [PageNumber.CURRENT], size: 18 }, { text: ' 页', size: 18 }],
          { align: AlignmentType.CENTER }
        )],
      }),
    },
    children: bodyChildren,
  };

  const coverPageMargin = {
    top: baseMargin,
    right: baseMargin,
    bottom: baseMargin,
    left: baseMargin,
  };

  const sections = coverTemplate
    ? [bodySection]
    : [
        {
          properties: {
            page: {
              size: { width: pageWidth, height: pageHeight },
              margin: coverPageMargin,
            },
          },
          children: coverChildren,
        },
        bodySection,
      ];

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font, size: 24 },
        },
      },
    },
    sections,
  });

  const outPath = config.output_path || '报告.docx';
  const finalPath = path.resolve(configDir, outPath);

  if (coverTemplate) {
    const templateFullPath = path.resolve(configDir, coverTemplate);
    if (!fs.existsSync(templateFullPath)) {
      console.error(`错误: 封面模板不存在: ${templateFullPath}`);
      process.exit(1);
    }

    const tmpPath = finalPath + '.tmp.docx';
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(tmpPath, buffer);

    const merged = mergeCoverTemplate(templateFullPath, tmpPath, finalPath);
    if (merged) {
      fs.unlinkSync(tmpPath);
      const stats = fs.statSync(finalPath);
      console.log(`报告已生成: ${outPath} (${(stats.size / 1024).toFixed(1)} KB) [使用封面模板]`);
    } else {
      fs.renameSync(tmpPath, finalPath);
      const stats = fs.statSync(finalPath);
      console.log(`报告已生成: ${outPath} (${(stats.size / 1024).toFixed(1)} KB) [封面模板合并失败，使用自动生成封面]`);
    }
  } else {
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(finalPath, buffer);
    console.log(`报告已生成: ${outPath} (${(buffer.length / 1024).toFixed(1)} KB)`);
  }
}

// ── CLI 直接运行 ────────────────────────────────────
if (require.main === module) {
  const configPath = process.argv[2];
  if (!configPath) {
    console.error('用法: node generate-report.js <config.json>');
    process.exit(1);
  }
  generateReport(configPath).catch(err => {
    console.error('生成报告失败:', err);
    process.exit(1);
  });
}

module.exports = { generateReport };
