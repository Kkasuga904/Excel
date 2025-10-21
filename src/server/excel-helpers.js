/**
 * Excel操作ヘルパー関数
 * セルデータの処理や分析を行う
 */

/**
 * セルデータから統計情報を計算
 */
function calculateStatistics(values) {
  if (!values || values.length === 0) {
    return null;
  }

  // 数値のみを抽出
  const numbers = [];
  values.forEach((row) => {
    if (Array.isArray(row)) {
      row.forEach((cell) => {
        const num = parseFloat(cell);
        if (!isNaN(num)) {
          numbers.push(num);
        }
      });
    } else {
      const num = parseFloat(row);
      if (!isNaN(num)) {
        numbers.push(num);
      }
    }
  });

  if (numbers.length === 0) {
    return null;
  }

  const sum = numbers.reduce((a, b) => a + b, 0);
  const average = sum / numbers.length;
  const max = Math.max(...numbers);
  const min = Math.min(...numbers);
  const median =
    numbers.length % 2 === 0
      ? (numbers[numbers.length / 2 - 1] + numbers[numbers.length / 2]) / 2
      : numbers[Math.floor(numbers.length / 2)];

  return {
    sum: sum.toFixed(2),
    average: average.toFixed(2),
    max: max.toFixed(2),
    min: min.toFixed(2),
    median: median.toFixed(2),
    count: numbers.length
  };
}

/**
 * 異常値（外れ値）を検出
 */
function detectOutliers(values) {
  const numbers = [];
  values.forEach((row) => {
    if (Array.isArray(row)) {
      row.forEach((cell) => {
        const num = parseFloat(cell);
        if (!isNaN(num)) {
          numbers.push(num);
        }
      });
    } else {
      const num = parseFloat(row);
      if (!isNaN(num)) {
        numbers.push(num);
      }
    }
  });

  if (numbers.length < 4) {
    return [];
  }

  // 四分位数を計算
  numbers.sort((a, b) => a - b);
  const q1 = numbers[Math.floor(numbers.length / 4)];
  const q3 = numbers[Math.floor((numbers.length * 3) / 4)];
  const iqr = q3 - q1;

  // IQR法で外れ値を検出
  const outliers = numbers.filter(
    (num) => num < q1 - 1.5 * iqr || num > q3 + 1.5 * iqr
  );

  return outliers;
}

/**
 * 前月比を計算
 */
function calculateMonthOverMonth(currentValues, previousValues) {
  if (!currentValues || !previousValues) {
    return null;
  }

  const current = parseFloat(currentValues);
  const previous = parseFloat(previousValues);

  if (isNaN(current) || isNaN(previous) || previous === 0) {
    return null;
  }

  const change = current - previous;
  const percentChange = ((change / previous) * 100).toFixed(2);

  return {
    current: current.toFixed(2),
    previous: previous.toFixed(2),
    change: change.toFixed(2),
    percentChange: percentChange,
    trend: change > 0 ? '増加' : change < 0 ? '減少' : '変化なし'
  };
}

/**
 * 重複を検出
 */
function detectDuplicates(values) {
  const seen = new Set();
  const duplicates = [];

  values.forEach((row, rowIndex) => {
    const rowString = JSON.stringify(row);
    if (seen.has(rowString)) {
      duplicates.push(rowIndex);
    } else {
      seen.add(rowString);
    }
  });

  return duplicates;
}

/**
 * 空白セルを補完（線形補間）
 */
function fillBlankCells(values) {
  if (!values || values.length === 0) {
    return values;
  }

  const filled = values.map((row) => [...row]);

  // 各列に対して処理
  for (let col = 0; col < filled[0].length; col++) {
    const column = filled.map((row) => row[col]);

    // 数値列を検出
    const numbers = column.map((cell) => parseFloat(cell));
    if (numbers.every((n) => isNaN(n))) {
      continue; // 数値列でない場合はスキップ
    }

    // 空白を補完
    for (let i = 0; i < column.length; i++) {
      if (column[i] === null || column[i] === '' || isNaN(numbers[i])) {
        // 前後の値を探す
        let prevIndex = -1;
        let nextIndex = -1;

        for (let j = i - 1; j >= 0; j--) {
          if (!isNaN(numbers[j])) {
            prevIndex = j;
            break;
          }
        }

        for (let j = i + 1; j < column.length; j++) {
          if (!isNaN(numbers[j])) {
            nextIndex = j;
            break;
          }
        }

        // 線形補間
        if (prevIndex !== -1 && nextIndex !== -1) {
          const prevValue = numbers[prevIndex];
          const nextValue = numbers[nextIndex];
          const interpolated =
            prevValue +
            ((nextValue - prevValue) / (nextIndex - prevIndex)) * (i - prevIndex);
          filled[i][col] = interpolated.toFixed(2);
        } else if (prevIndex !== -1) {
          filled[i][col] = numbers[prevIndex];
        } else if (nextIndex !== -1) {
          filled[i][col] = numbers[nextIndex];
        }
      }
    }
  }

  return filled;
}

/**
 * データをソート
 */
function sortData(values, columnIndex, ascending = true) {
  if (!values || values.length === 0) {
    return values;
  }

  const sorted = [...values].sort((a, b) => {
    const aVal = a[columnIndex];
    const bVal = b[columnIndex];

    const aNum = parseFloat(aVal);
    const bNum = parseFloat(bVal);

    if (!isNaN(aNum) && !isNaN(bNum)) {
      return ascending ? aNum - bNum : bNum - aNum;
    }

    const aStr = String(aVal).toLowerCase();
    const bStr = String(bVal).toLowerCase();
    return ascending ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
  });

  return sorted;
}

/**
 * 月次レポートを生成
 */
function generateMonthlyReport(values, address) {
  const stats = calculateStatistics(values);
  if (!stats) {
    return null;
  }

  const report = [
    ['月次レポート'],
    [''],
    ['データ範囲', address],
    [''],
    ['統計情報'],
    ['合計', stats.sum],
    ['平均', stats.average],
    ['最大値', stats.max],
    ['最小値', stats.min],
    ['中央値', stats.median],
    ['データ数', stats.count]
  ];

  return report;
}

module.exports = {
  calculateStatistics,
  detectOutliers,
  calculateMonthOverMonth,
  detectDuplicates,
  fillBlankCells,
  sortData,
  generateMonthlyReport
};

