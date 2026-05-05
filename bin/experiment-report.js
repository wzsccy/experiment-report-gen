#!/usr/bin/env node

const path = require('path');

const usage = `
experiment-report-gen - 深度学习实验报告生成器

用法:
  experiment-report generate <config.json>   从配置文件生成 .docx 报告
  experiment-report analyze <results.json>   从 results.json 自动生成分析文本
  experiment-report --help                   显示帮助信息

示例:
  experiment-report generate config.json
  experiment-report analyze output/results.json --type image
  experiment-report analyze output/results.json --type text
`.trim();

const args = process.argv.slice(2);
const command = args[0];

if (!command || command === '--help' || command === '-h') {
  console.log(usage);
  process.exit(0);
}

if (command === 'generate') {
  const configPath = args[1];
  if (!configPath) {
    console.error('错误: 请指定 config.json 路径');
    console.error('用法: experiment-report generate <config.json>');
    process.exit(1);
  }
  const { generateReport } = require('../src/generate-report');
  generateReport(path.resolve(configPath)).catch(err => {
    console.error('生成报告失败:', err.message || err);
    process.exit(1);
  });
} else if (command === 'analyze') {
  const resultsPath = args[1];
  if (!resultsPath) {
    console.error('错误: 请指定 results.json 路径');
    console.error('用法: experiment-report analyze <results.json> [--type image|text|auto]');
    process.exit(1);
  }
  let type = 'auto';
  const typeIdx = args.indexOf('--type');
  if (typeIdx !== -1 && args[typeIdx + 1]) {
    type = args[typeIdx + 1];
  }
  const { analyzeResults } = require('../src/analyze-results');
  try {
    const result = analyzeResults(path.resolve(resultsPath), type);
    console.log(result);
  } catch (err) {
    console.error('分析失败:', err.message || err);
    process.exit(1);
  }
} else {
  console.error(`未知命令: ${command}`);
  console.error(usage);
  process.exit(1);
}
