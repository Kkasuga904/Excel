/**
 * 高度なExcel操作ヘルパー
 * グラフ生成、複雑な分析など
 */

/**
 * グラフ生成用のデータを準備
 */
function prepareChartData(values, chartType = 'line') {
  if (!values || values.length === 0) {
    return null;
  }

  // 最初の行をラベルとして使用
  const labels = values[0];
  const datasets = [];

  // 2行目以降をデータとして使用
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const dataset = {
      label: row[0] || `Series ${i}`,
      data: row.slice(1).map((v) => parseFloat(v) || 0)
    };
    datasets.push(dataset);
  }

  return {
    labels: labels.slice(1),
    datasets: datasets,
    chartType: chartType
  };
}

/**
 * ピボットテーブルを生成
 */
function generatePivotTable(values, rowIndex, colIndex, valueIndex) {
  if (!values || values.length === 0) {
    return null;
  }

  const pivot = {};

  // ピボットテーブルを構築
  values.forEach((row) => {
    const rowKey = row[rowIndex];
    const colKey = row[colIndex];
    const value = parseFloat(row[valueIndex]) || 0;

    if (!pivot[rowKey]) {
      pivot[rowKey] = {};
    }

    if (!pivot[rowKey][colKey]) {
      pivot[rowKey][colKey] = 0;
    }

    pivot[rowKey][colKey] += value;
  });

  // 2次元配列に変換
  const result = [];
  const columns = new Set();

  // 列を収集
  Object.values(pivot).forEach((row) => {
    Object.keys(row).forEach((col) => columns.add(col));
  });

  // ヘッダー行
  result.push(['', ...Array.from(columns)]);

  // データ行
  Object.entries(pivot).forEach(([rowKey, rowData]) => {
    const row = [rowKey];
    Array.from(columns).forEach((col) => {
      row.push(rowData[col] || 0);
    });
    result.push(row);
  });

  return result;
}

/**
 * 条件付き書式用のデータを生成
 */
function generateConditionalFormatting(values) {
  if (!values || values.length === 0) {
    return null;
  }

  const numbers = [];
  const positions = [];

  // 数値を抽出
  values.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const num = parseFloat(cell);
      if (!isNaN(num)) {
        numbers.push(num);
        positions.push({ row: rowIndex, col: colIndex, value: num });
      }
    });
  });

  if (numbers.length === 0) {
    return null;
  }

  // 最大値と最小値を計算
  const max = Math.max(...numbers);
  const min = Math.min(...numbers);
  const range = max - min || 1;

  // 各セルのカラースケールを計算
  const formatting = positions.map((pos) => {
    const normalized = (pos.value - min) / range;
    const color = getColorForValue(normalized);
    return {
      address: `${String.fromCharCode(65 + pos.col)}${pos.row + 1}`,
      color: color,
      value: pos.value
    };
  });

  return formatting;
}

/**
 * 正規化された値（0-1）から色を取得
 */
function getColorForValue(normalized) {
  // 赤（低）から緑（高）へのグラデーション
  const hue = normalized * 120; // 0°（赤）から120°（緑）
  return `hsl(${hue}, 100%, 50%)`;
}

/**
 * 回帰分析を実行
 */
function performRegression(xValues, yValues) {
  if (!xValues || !yValues || xValues.length !== yValues.length) {
    return null;
  }

  const n = xValues.length;
  const sumX = xValues.reduce((a, b) => a + b, 0);
  const sumY = yValues.reduce((a, b) => a + b, 0);
  const sumXY = xValues.reduce((sum, x, i) => sum + x * yValues[i], 0);
  const sumX2 = xValues.reduce((sum, x) => sum + x * x, 0);

  // 最小二乗法で係数を計算
  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;

  // R二乗を計算
  const yMean = sumY / n;
  const ssTotal = yValues.reduce((sum, y) => sum + (y - yMean) ** 2, 0);
  const ssResidual = yValues.reduce((sum, y, i) => {
    const predicted = slope * xValues[i] + intercept;
    return sum + (y - predicted) ** 2;
  }, 0);
  const rSquared = 1 - ssResidual / ssTotal;

  return {
    slope: slope.toFixed(4),
    intercept: intercept.toFixed(4),
    rSquared: rSquared.toFixed(4),
    equation: `y = ${slope.toFixed(4)}x + ${intercept.toFixed(4)}`
  };
}

/**
 * 相関係数を計算
 */
function calculateCorrelation(series1, series2) {
  if (!series1 || !series2 || series1.length !== series2.length) {
    return null;
  }

  const n = series1.length;
  const mean1 = series1.reduce((a, b) => a + b, 0) / n;
  const mean2 = series2.reduce((a, b) => a + b, 0) / n;

  const numerator = series1.reduce((sum, x, i) => {
    return sum + (x - mean1) * (series2[i] - mean2);
  }, 0);

  const denominator = Math.sqrt(
    series1.reduce((sum, x) => sum + (x - mean1) ** 2, 0) *
    series2.reduce((sum, y) => sum + (y - mean2) ** 2, 0)
  );

  if (denominator === 0) {
    return null;
  }

  return (numerator / denominator).toFixed(4);
}

/**
 * 時系列データの移動平均を計算
 */
function calculateMovingAverage(values, period = 3) {
  if (!values || values.length < period) {
    return null;
  }

  const result = [];

  for (let i = 0; i < values.length; i++) {
    if (i < period - 1) {
      result.push(null);
    } else {
      const sum = values
        .slice(i - period + 1, i + 1)
        .reduce((a, b) => a + parseFloat(b), 0);
      result.push((sum / period).toFixed(2));
    }
  }

  return result;
}

/**
 * データの正規化（0-1スケール）
 */
function normalizeData(values) {
  if (!values || values.length === 0) {
    return null;
  }

  const numbers = values.map((v) => parseFloat(v)).filter((v) => !isNaN(v));

  if (numbers.length === 0) {
    return null;
  }

  const min = Math.min(...numbers);
  const max = Math.max(...numbers);
  const range = max - min || 1;

  return numbers.map((v) => (((v - min) / range).toFixed(4)));
}

/**
 * Z-スコアを計算（異常値検出）
 */
function calculateZScores(values) {
  if (!values || values.length === 0) {
    return null;
  }

  const numbers = values.map((v) => parseFloat(v)).filter((v) => !isNaN(v));

  if (numbers.length < 2) {
    return null;
  }

  const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
  const variance =
    numbers.reduce((sum, v) => sum + (v - mean) ** 2, 0) / numbers.length;
  const stdDev = Math.sqrt(variance);

  if (stdDev === 0) {
    return null;
  }

  return numbers.map((v) => {
    const zScore = ((v - mean) / stdDev).toFixed(4);
    return {
      value: v,
      zScore: zScore,
      isOutlier: Math.abs(zScore) > 3 // 3σ以上は外れ値
    };
  });
}

module.exports = {
  prepareChartData,
  generatePivotTable,
  generateConditionalFormatting,
  getColorForValue,
  performRegression,
  calculateCorrelation,
  calculateMovingAverage,
  normalizeData,
  calculateZScores
};

