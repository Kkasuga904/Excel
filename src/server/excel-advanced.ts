import { CellMatrix, ChartData, ChartType, PivotTableResult } from './types';

const toNumber = (value: unknown): number | null => {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === 'string') {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
};

export const prepareChartData = (values: CellMatrix, chartType: ChartType = 'LineChart'): ChartData | null => {
  if (!values || values.length < 2) {
    return null;
  }

  const [header, ...rows] = values;
  const labels = header.slice(1).map((label) => String(label ?? ''));
  const datasets = rows.map((row, index) => ({
    label: String(row[0] ?? `Series ${index + 1}`),
    data: row
      .slice(1)
      .map((cell) => toNumber(cell) ?? 0)
  }));

  return {
    labels,
    datasets,
    chartType
  };
};

export const generatePivotTable = (
  values: CellMatrix,
  rowIndex: number,
  colIndex: number,
  valueIndex: number
): PivotTableResult | null => {
  if (!values || values.length === 0) {
    return null;
  }

  const pivotMap = new Map<string, Map<string, number>>();
  const columnKeys = new Set<string>();

  values.forEach((row) => {
    const rowKey = String(row[rowIndex] ?? '');
    const colKey = String(row[colIndex] ?? '');
    const value = toNumber(row[valueIndex]) ?? 0;

    if (!pivotMap.has(rowKey)) {
      pivotMap.set(rowKey, new Map());
    }

    const columnMap = pivotMap.get(rowKey)!;
    columnMap.set(colKey, (columnMap.get(colKey) ?? 0) + value);
    columnKeys.add(colKey);
  });

  const orderedColumns = Array.from(columnKeys);
  const table: (string | number)[][] = [['', ...orderedColumns]];

  pivotMap.forEach((columnMap, rowKey) => {
    const row = [rowKey, ...orderedColumns.map((col) => columnMap.get(col) ?? 0)];
    table.push(row);
  });

  return {
    columns: orderedColumns,
    rows: table
  };
};

export const generateConditionalFormatting = (values: CellMatrix) => {
  if (!values || values.length === 0) {
    return null;
  }

  const points: { row: number; col: number; value: number }[] = [];

  values.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const parsed = toNumber(cell);
      if (parsed !== null) {
        points.push({ row: rowIndex, col: colIndex, value: parsed });
      }
    });
  });

  if (points.length === 0) {
    return null;
  }

  const valuesOnly = points.map((point) => point.value);
  const min = Math.min(...valuesOnly);
  const max = Math.max(...valuesOnly);
  const range = max - min || 1;

  return points.map((point) => {
    const normalized = (point.value - min) / range;
    return {
      address: `${String.fromCharCode(65 + point.col)}${point.row + 1}`,
      color: getColorForValue(normalized),
      value: point.value
    };
  });
};

export const getColorForValue = (normalized: number): string => {
  const bounded = Math.min(Math.max(normalized, 0), 1);
  const hue = bounded * 120;
  return `hsl(${hue.toFixed(0)}, 100%, 50%)`;
};

export const performRegression = (xValues: number[], yValues: number[]) => {
  if (!xValues.length || xValues.length !== yValues.length) {
    return null;
  }

  const n = xValues.length;
  const sumX = xValues.reduce((acc, value) => acc + value, 0);
  const sumY = yValues.reduce((acc, value) => acc + value, 0);
  const sumXY = xValues.reduce((acc, x, index) => acc + x * yValues[index], 0);
  const sumXSquare = xValues.reduce((acc, x) => acc + x ** 2, 0);

  const denominator = n * sumXSquare - sumX ** 2;
  if (denominator === 0) {
    return null;
  }

  const slope = (n * sumXY - sumX * sumY) / denominator;
  const intercept = (sumY - slope * sumX) / n;
  const meanY = sumY / n;
  const totalSS = yValues.reduce((acc, y) => acc + (y - meanY) ** 2, 0);
  const residualSS = yValues.reduce((acc, y, index) => {
    const predicted = slope * xValues[index] + intercept;
    return acc + (y - predicted) ** 2;
  }, 0);
  const rSquared = totalSS === 0 ? 1 : 1 - residualSS / totalSS;

  return {
    slope: slope.toFixed(4),
    intercept: intercept.toFixed(4),
    rSquared: rSquared.toFixed(4),
    equation: `y = ${slope.toFixed(4)}x + ${intercept.toFixed(4)}`
  };
};

export const calculateCorrelation = (series1: number[], series2: number[]) => {
  if (!series1.length || series1.length !== series2.length) {
    return null;
  }

  const n = series1.length;
  const mean1 = series1.reduce((acc, value) => acc + value, 0) / n;
  const mean2 = series2.reduce((acc, value) => acc + value, 0) / n;

  const numerator = series1.reduce((acc, value, index) => acc + (value - mean1) * (series2[index] - mean2), 0);
  const denominator = Math.sqrt(
    series1.reduce((acc, value) => acc + (value - mean1) ** 2, 0) *
      series2.reduce((acc, value) => acc + (value - mean2) ** 2, 0)
  );

  if (denominator === 0) {
    return null;
  }

  return (numerator / denominator).toFixed(4);
};

export const calculateMovingAverage = (series: number[], period = 3): (string | null)[] | null => {
  if (!series.length || series.length < period) {
    return null;
  }

  return series.map((_, index) => {
    if (index < period - 1) {
      return null;
    }
    const slice = series.slice(index - period + 1, index + 1);
    const sum = slice.reduce((acc, value) => acc + value, 0);
    return (sum / period).toFixed(2);
  });
};

export const normalizeData = (series: number[]): string[] | null => {
  if (!series.length) {
    return null;
  }

  const min = Math.min(...series);
  const max = Math.max(...series);
  const range = max - min || 1;

  return series.map((value) => ((value - min) / range).toFixed(4));
};

export const calculateZScores = (series: number[]) => {
  if (!series.length) {
    return null;
  }

  const mean = series.reduce((acc, value) => acc + value, 0) / series.length;
  const variance = series.reduce((acc, value) => acc + (value - mean) ** 2, 0) / series.length;
  const stdDev = Math.sqrt(variance);

  if (stdDev === 0) {
    return null;
  }

  return series.map((value) => {
    const zScore = (value - mean) / stdDev;
    return {
      value,
      zScore: zScore.toFixed(4),
      isOutlier: Math.abs(zScore) > 3
    };
  });
};
