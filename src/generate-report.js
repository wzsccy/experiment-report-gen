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
const SINGLE_BORDER = { style: BorderStyle.SINGLE, size: 1, color: '000000' };
const SINGLE_BORDERS = { top: SINGLE_BORDER, bottom: SINGLE_BORDER, left: SINGLE_BORDER, right: SINGLE_BORDER };

const FONT = { ascii: 'Times New Roman', eastAsia: '宋体' };

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

function textToParagraphs(text, fontSize = 24) {
  const result = [];
  const blocks = text.split('\n\n');
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

function explainParagraph(text) {
  if (!text) return [];
  return [new Paragraph({
    spacing: { after: 200, line: 360 },
    indent: { firstLine: 480 },
    children: [new TextRun({ text, font: FONT, size: 24 })],
  })];
}

// ── 数据表格（无蓝色表头，居中）─────────────────────

function createDataTable(tableConfig) {
  const { caption, headers, rows, explain } = tableConfig;
  const paragraphs = [];

  if (caption) {
    paragraphs.push(new Paragraph({
      spacing: { before: 200, after: 100 },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: caption, bold: true, font: FONT, size: 21 })],
    }));
  }

  const colCount = headers.length;
  const colWidth = Math.floor(9000 / colCount);
  const colWidths = Array(colCount).fill(colWidth);
  colWidths[colCount - 1] = 9000 - colWidth * (colCount - 1);

  // 表头行（无蓝色背景，仅加粗）
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, idx) => new TableCell({
      borders: SINGLE_BORDERS,
      width: { size: colWidths[idx], type: WidthType.DXA },
      verticalAlign: 'center',
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
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
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 0 },
          children: [new TextRun({ text: String(cell), font: FONT, size: 21 })],
        })],
      })),
    })
  );

  const table = new Table({
    width: { size: 9000, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });

  paragraphs.push(table);
  paragraphs.push(new Paragraph({ spacing: { after: 100 }, children: [] }));

  // 解释文字
  paragraphs.push(...explainParagraph(explain));

  return paragraphs;
}

// ── 图片加载（支持 explain 字段）─────────────────────

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
  paragraphs.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
  paragraphs.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
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
      size: 32, bold: true, align: AlignmentType.CENTER, spaceBefore: 200, spaceAfter: 200
    }));
  }

  if (config.date) {
    children.push(p(`日期：${config.date}`, { size: 24, spaceAfter: 200 }));
  }

  const sectionRows = bodySections.map((section, i) => {
    const key = section.key || `section_${i}`;
    const cellChildren = [];

    const sectionTitle = section.title || `第${i + 1}节`;
    cellChildren.push(new Paragraph({
      spacing: { after: 120, line: 360 },
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

    return new TableRow({
      children: [
        new TableCell({
          borders: SINGLE_BORDERS,
          width: { size: 9000, type: WidthType.DXA },
          margins: { top: 100, bottom: 100, left: 150, right: 150 },
          children: cellChildren,
        }),
      ],
    });
  });

  if (sectionRows.length > 0) {
    const mainTable = new Table({
      width: { size: 9000, type: WidthType.DXA },
      columnWidths: [9000],
      rows: sectionRows,
    });
    children.push(mainTable);
  }

  return children;
}

// ── 实验报告模式: 构建正文 (向后兼容) ──────────────
function buildExperimentBody(config, sectionImages, sectionTables) {
  const { title, date, sections } = config;
  const children = [];

  children.push(p(title, {
    size: 32, bold: true, align: AlignmentType.CENTER, spaceBefore: 200, spaceAfter: 200
  }));
  children.push(p(`实验时间：${date || ''}`, { size: 24, spaceAfter: 200 }));

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

    cellChildren.push(new Paragraph({
      spacing: { after: 120, line: 360 },
      children: [new TextRun({ text: titleText, bold: true, font: FONT, size: 24 })],
    }));

    const tablesForSection = sectionTables[key] || [];
    for (const tbl of tablesForSection) {
      if (tbl.position === 'start') {
        cellChildren.push(...createDataTable(tbl));
      }
    }

    cellChildren.push(...textToParagraphs(content));

    for (const tbl of tablesForSection) {
      if (tbl.position !== 'start') {
        cellChildren.push(...createDataTable(tbl));
      }
    }

    if (sectionImages[key] && sectionImages[key].length > 0) {
      cellChildren.push(...sectionImages[key]);
    }

    return new TableRow({
      children: [
        new TableCell({
          borders: SINGLE_BORDERS,
          width: { size: 9000, type: WidthType.DXA },
          margins: { top: 100, bottom: 100, left: 150, right: 150 },
          children: cellChildren,
        }),
      ],
    });
  });

  const mainTable = new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [9000],
    rows: sectionRows,
  });
  children.push(mainTable);

  // 思考题
  children.push(emptyLine());
  const thoughtCellChildren = [];
  thoughtCellChildren.push(new Paragraph({
    spacing: { after: 120, line: 360 },
    children: [new TextRun({ text: '相关思考题及解答', bold: true, font: FONT, size: 24 })],
  }));
  if (sections.thought_questions) {
    thoughtCellChildren.push(...textToParagraphs(sections.thought_questions));
  }

  const thoughtTable = new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [9000],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: SINGLE_BORDERS,
            width: { size: 9000, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            children: thoughtCellChildren,
          }),
        ],
      }),
    ],
  });
  children.push(thoughtTable);

  return children;
}

// ── 封面模板合并（调用 Python 脚本）──────────────────

function mergeCoverTemplate(templatePath, bodyDocxPath, outputPath) {
  // 查找 merge-cover.py 脚本
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
    // 尝试 python（Windows）
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

/**
 * 从 config 生成 .docx 报告
 *
 * 支持两种模式:
 *   - mode: "generic" — 完全自定义封面和章节
 *   - 默认 (experiment) — 向后兼容实验报告格式
 *
 * @param {string} configPath - config.json 的路径
 * @returns {Promise<void>}
 */
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
  const margin = config.margin || 1440;

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font, size: 24 },
        },
      },
    },
    sections: [
      // 封面
      {
        properties: {
          page: {
            size: { width: pageWidth, height: pageHeight },
            margin: { top: margin, right: margin, bottom: margin, left: margin },
          },
        },
        children: coverChildren,
      },
      // 正文
      {
        properties: {
          page: {
            size: { width: pageWidth, height: pageHeight },
            margin: { top: margin, right: margin, bottom: margin, left: margin },
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
      },
    ],
  });

  const outPath = config.output_path || '报告.docx';
  const finalPath = path.resolve(configDir, outPath);

  if (coverTemplate) {
    // 有封面模板：先生成临时文件，再合并
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
      // 合并失败，使用原始文件
      fs.renameSync(tmpPath, finalPath);
      const stats = fs.statSync(finalPath);
      console.log(`报告已生成: ${outPath} (${(stats.size / 1024).toFixed(1)} KB) [封面模板合并失败，使用自动生成封面]`);
    }
  } else {
    // 无封面模板：直接生成
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
