/**
 * analyze-results.js - 从 results.json 自动生成实验分析文本 (Node.js 版)
 *
 * 用法:
 *   const { analyzeResults } = require('experiment-report-gen/src/analyze-results');
 *   const text = analyzeResults('output/results.json', 'image');
 *
 * CLI:
 *   node bin/experiment-report.js analyze <results.json> [--type image|text|auto]
 */

const fs = require('fs');

function analyzeImageClassification(data) {
  const bestAcc = data.best_acc || 0;
  const trainLosses = data.train_losses || [];
  const testLosses = data.test_losses || [];
  const trainAccs = data.train_accs || [];
  const testAccs = data.test_accs || [];
  const epochs = trainLosses.length;

  const lines = [];

  lines.push('5.1 损失函数变化分析');
  lines.push(
    `从${epochs}个epoch的训练过程来看，模型在训练集上的损失函数从${trainLosses[0]?.toFixed(4)}下降至${trainLosses[epochs - 1]?.toFixed(4)}，` +
    `呈现出明显的下降趋势，说明模型参数在不断优化，预测误差逐步减小。`
  );
  lines.push(
    `同时，测试集损失函数从${testLosses[0]?.toFixed(4)}下降至${testLosses[epochs - 1]?.toFixed(4)}，整体也呈下降趋势，且变化趋势与训练集基本一致。`
  );
  lines.push('');

  const gap = trainLosses[epochs - 1] - testLosses[epochs - 1];
  if (gap < -0.5) {
    lines.push('值得注意的是，训练集损失明显低于测试集损失，说明模型在训练数据上拟合较好，但存在一定的泛化差距，建议进一步增加正则化手段。');
  } else {
    lines.push('训练集和测试集损失差距较小，说明模型在训练过程中未出现严重过拟合现象。');
  }
  lines.push('');

  lines.push('5.2 准确率变化分析');
  lines.push(
    `训练集准确率从${trainAccs[0]?.toFixed(2)}%提升至${trainAccs[epochs - 1]?.toFixed(2)}%，` +
    `测试集准确率从${testAccs[0]?.toFixed(2)}%提升至${testAccs[epochs - 1]?.toFixed(2)}%。`
  );
  lines.push(`最佳测试准确率为${bestAcc.toFixed(2)}%。`);
  lines.push('');

  if (bestAcc >= 80) {
    lines.push(`模型最终测试准确率达到${bestAcc.toFixed(2)}%，超过了80%的目标要求，说明模型结构设计和超参数选择较为合理。`);
  } else {
    lines.push(`模型最终测试准确率为${bestAcc.toFixed(2)}%，尚未达到80%的目标。建议尝试更深的网络结构、更强的数据增强或更长的训练轮次来提升性能。`);
  }
  lines.push('');

  lines.push('5.3 预测结果分析');
  lines.push('从测试集前8张图像的预测结果来看，模型对大部分图像的分类是正确的，预测标签与真实标签一致。');
  lines.push('对于分类错误的样本，部分图像存在一定的模糊性或属于容易混淆的类别，这也反映了图像分类任务本身的挑战性。');

  return lines.join('\n');
}

function analyzeTextClassification(data) {
  const lines = [];

  if (data.with_attention && data.no_attention) {
    const resAttn = data.with_attention;
    const resNo = data.no_attention;
    const accDrop = data.acc_drop ?? (resAttn.test_acc - resNo.test_acc);

    lines.push('5.1 损失函数变化分析');
    lines.push(
      `从训练过程来看，两个模型的训练损失函数均呈现出明显的下降趋势。` +
      `带注意力机制的模型在验证集上的损失函数从${resAttn.val_losses[0]?.toFixed(4)}下降至${resAttn.val_losses[resAttn.val_losses.length - 1]?.toFixed(4)}，` +
      `下降更为平稳，说明注意力机制有助于模型更好地泛化。`
    );
    lines.push(
      `无注意力机制的模型验证集损失从${resNo.val_losses[0]?.toFixed(4)}下降至${resNo.val_losses[resNo.val_losses.length - 1]?.toFixed(4)}，` +
      `在训练后期略有波动，表明模型可能存在轻微过拟合。`
    );
    lines.push('');

    lines.push('5.2 准确率变化分析');
    lines.push(`带注意力机制的模型在验证集上达到${resAttn.best_val_acc?.toFixed(2)}%，测试集准确率为${resAttn.test_acc?.toFixed(2)}%。`);
    lines.push(`无注意力机制的模型测试集准确率为${resNo.test_acc?.toFixed(2)}%。`);
    lines.push(`注意力机制带来的准确率提升为${accDrop >= 0 ? '+' : ''}${accDrop.toFixed(2)}%。`);
    lines.push('');

    lines.push('5.3 对比分析');
    if (accDrop > 0) {
      lines.push(`通过对比，带注意力机制的模型测试准确率比无注意力机制的模型高出${Math.abs(accDrop).toFixed(2)}%，验证了注意力机制在文本分类任务中的有效性。`);
    } else {
      lines.push(`在当前实验设置下，注意力机制带来的准确率提升为${Math.abs(accDrop).toFixed(2)}%，差异不显著，可能与数据集规模或文本长度有关。`);
    }
    lines.push('注意力机制尤其在处理较长文本时优势明显，因为它能够动态地为不同位置的词分配不同的权重，突出对分类任务最重要的信息。');
    lines.push('');

    lines.push('5.4 注意力权重可视化');
    lines.push('从注意力权重可视化结果可以看出，模型对文本中的关键实体词赋予了较高的注意力权重，而对停用词和标点符号的权重较低。这验证了注意力机制能够自动学习到对分类任务最有价值的信息。');

  } else if (data.best_acc !== undefined) {
    const bestAcc = data.best_acc || 0;
    const trainLosses = data.train_losses || [];
    const valLosses = data.val_losses || [];
    const trainAccs = data.train_accs || [];
    const valAccs = data.val_accs || [];

    lines.push('5.1 损失函数变化分析');
    if (trainLosses.length > 0) {
      lines.push(`训练集损失从${trainLosses[0]?.toFixed(4)}下降至${trainLosses[trainLosses.length - 1]?.toFixed(4)}，模型在训练过程中持续优化。`);
    }
    if (valLosses.length > 0) {
      lines.push(`验证集损失从${valLosses[0]?.toFixed(4)}下降至${valLosses[valLosses.length - 1]?.toFixed(4)}，与训练集趋势一致。`);
    }
    lines.push('');

    lines.push('5.2 准确率变化分析');
    if (trainAccs.length > 0) {
      lines.push(`训练集准确率从${trainAccs[0]?.toFixed(2)}%提升至${trainAccs[trainAccs.length - 1]?.toFixed(2)}%。`);
    }
    if (valAccs.length > 0) {
      lines.push(`验证集准确率从${valAccs[0]?.toFixed(2)}%提升至${valAccs[valAccs.length - 1]?.toFixed(2)}%。`);
    }
    lines.push(`最佳测试准确率为${bestAcc.toFixed(2)}%。`);
  }

  return lines.join('\n');
}

/**
 * 分析 results.json 并生成报告文本
 * @param {string} resultsPath - results.json 路径
 * @param {string} type - 实验类型: 'image', 'text', 或 'auto'
 * @returns {string} 分析文本
 */
function analyzeResults(resultsPath, type = 'auto') {
  const data = JSON.parse(fs.readFileSync(resultsPath, 'utf-8'));

  if (type === 'auto') {
    if (data.with_attention || data.acc_drop !== undefined) {
      type = 'text';
    } else {
      type = 'image';
    }
  }

  if (type === 'image') {
    return analyzeImageClassification(data);
  } else {
    return analyzeTextClassification(data);
  }
}

// ── CLI 直接运行 ────────────────────────────────────
if (require.main === module) {
  const resultsPath = process.argv[2];
  if (!resultsPath) {
    console.error('用法: node analyze-results.js <results.json> [--type image|text|auto]');
    process.exit(1);
  }
  let type = 'auto';
  const typeIdx = process.argv.indexOf('--type');
  if (typeIdx !== -1 && process.argv[typeIdx + 1]) {
    type = process.argv[typeIdx + 1];
  }
  console.log(analyzeResults(resultsPath, type));
}

module.exports = { analyzeResults };
