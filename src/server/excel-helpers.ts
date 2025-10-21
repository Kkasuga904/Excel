import { CellData, CellMatrix, CellValue, StatisticsSummary } from './types';

const isNumeric = (value: CellValue): value is number => {
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value === 'number') {
    return Number.isFinite(value);
  }
  if (typeof value === 'string') {
    const parsed = Number(value);
    return Number.isFinite(parsed);
  }
  return false;
};

const toNumber = (value: CellValue): number | null => {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
};

const flattenToNumbers = (values: CellMatrix | CellValue[]): number[] => {
  const numbers: number[] = [];
  values.forEach((row) => {
    if (Array.isArray(row)) {
      row.forEach((cell) => {
        const parsed = toNumber(cell);
        if (parsed !== null) {
          numbers.push(parsed);
        }
      });
    } else {
      const parsed = toNumber(row);
      if (parsed !== null) {
        numbers.push(parsed);
      }
    }
  });
  return numbers;
};

const format = (value: number): string => value.toFixed(2);

export const calculateStatistics = (values: CellMatrix): StatisticsSummary | null => {
  if (!values || values.length === 0) {
    return null;
  }

  const numbers = flattenToNumbers(values);
  if (numbers.length === 0) {
    return null;
  }

  const sorted = [...numbers].sort((a, b) => a - b);
  const sum = sorted.reduce((acc, num) => acc + num, 0);
  const average = sum / sorted.length;
  const midpoint = sorted.length / 2;
  const median =
    sorted.length % 2 === 0
      ? (sorted[midpoint - 1] + sorted[midpoint]) / 2
      : sorted[Math.floor(midpoint)];

  return {
    sum: format(sum),
    average: format(average),
    max: format(sorted[sorted.length - 1]),
    min: format(sorted[0]),
    median: format(median),
    count: sorted.length
  };
};

export const detectOutliers = (values: CellMatrix): number[] => {
  const numbers = flattenToNumbers(values);
  if (numbers.length < 4) {
    return [];
  }

  const sorted = [...numbers].sort((a, b) => a - b);
  const q1 = sorted[Math.floor(sorted.length / 4)];
  const q3 = sorted[Math.floor((sorted.length * 3) / 4)];
  const iqr = q3 - q1;

  return sorted.filter((value) => value < q1 - 1.5 * iqr || value > q3 + 1.5 * iqr);
};

export const calculateMonthOverMonth = (
  current: CellValue,
  previous: CellValue
): {
  current: string;
  previous: string;
  change: string;
  percentChange: string;
  trend: '増加' | '減少' | '変化なし';
} | null => {
  const currentValue = toNumber(current);
  const previousValue = toNumber(previous);

  if (
    currentValue === null ||
    previousValue === null ||
    Number.isNaN(currentValue) ||
    Number.isNaN(previousValue) ||
    previousValue === 0
  ) {
    return null;
  }

  const change = currentValue - previousValue;
  const percent = (change / previousValue) * 100;

  return {
    current: format(currentValue),
    previous: format(previousValue),
    change: format(change),
    percentChange: percent.toFixed(2),
    trend: change > 0 ? '増加' : change < 0 ? '減少' : '変化なし'
  };
};

export const detectDuplicates = (values: CellMatrix): number[] => {
  const seen = new Map<string, number>();
  const duplicates: number[] = [];

  values.forEach((row, index) => {
    const signature = JSON.stringify(row ?? []);
    if (seen.has(signature)) {
      duplicates.push(index);
    } else {
      seen.set(signature, index);
    }
  });

  return duplicates;
};

export const fillBlankCells = (values: CellMatrix): CellMatrix => {
  if (!values || values.length === 0) {
    return values;
  }

  const filled = values.map((row) => [...row]);
  const columnCount = filled[0]?.length ?? 0;

  for (let col = 0; col < columnCount; col += 1) {
    const columnValues = filled.map((row) => row[col]);
    const numericColumn = columnValues.map((cell) => toNumber(cell));

    if (numericColumn.every((value) => value === null)) {
      continue;
    }

    for (let rowIndex = 0; rowIndex < columnValues.length; rowIndex += 1) {
      if (numericColumn[rowIndex] !== null) {
        continue;
      }

      let prevIndex = rowIndex - 1;
      while (prevIndex >= 0 && numericColumn[prevIndex] === null) {
        prevIndex -= 1;
      }

      let nextIndex = rowIndex + 1;
      while (nextIndex < columnValues.length && numericColumn[nextIndex] === null) {
        nextIndex += 1;
      }

      if (prevIndex >= 0 && nextIndex < columnValues.length) {
        const prevValue = numericColumn[prevIndex] as number;
        const nextValue = numericColumn[nextIndex] as number;
        const interpolated =
          prevValue + ((nextValue - prevValue) / (nextIndex - prevIndex)) * (rowIndex - prevIndex);
        filled[rowIndex][col] = format(interpolated);
      } else if (prevIndex >= 0) {
        filled[rowIndex][col] = format(numericColumn[prevIndex] as number);
      } else if (nextIndex < columnValues.length) {
        filled[rowIndex][col] = format(numericColumn[nextIndex] as number);
      }
    }
  }

  return filled;
};

export const sortData = (values: CellMatrix, columnIndex: number, ascending = true): CellMatrix => {
  if (!values || values.length === 0) {
    return values;
  }

  const sorted = [...values].sort((rowA, rowB) => {
    const aVal = rowA[columnIndex];
    const bVal = rowB[columnIndex];
    const aNum = toNumber(aVal);
    const bNum = toNumber(bVal);

    if (aNum !== null && bNum !== null) {
      return ascending ? aNum - bNum : bNum - aNum;
    }

    const aStr = String(aVal ?? '').toLowerCase();
    const bStr = String(bVal ?? '').toLowerCase();
    return ascending ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
  });

  return sorted;
};

export const generateMonthlyReport = (values: CellMatrix, cellData?: CellData): string[][] | null => {
  const stats = calculateStatistics(values);
  if (!stats) {
    return null;
  }

  return [
    ['月次レポート'],
    [''],
    ['データ範囲', cellData?.address ?? '不明'],
    [''],
    ['統計情報'],
    ['合計', stats.sum],
    ['平均', stats.average],
    ['最大値', stats.max],
    ['最小値', stats.min],
    ['中央値', stats.median],
    ['データ数', String(stats.count)]
  ];
};

export const sanitizeCellData = (cellData?: CellData | null): CellData | null => {
  if (!cellData) {
    return null;
  }

  const safeValues = cellData.values.map((row) =>
    row.map((value) => {
      if (value === null || value === undefined) {
        return null;
      }
      if (typeof value === 'number') {
        return Number.isFinite(value) ? value : null;
      }
      return String(value);
    })
  );

  return {
    address: cellData.address,
    values: safeValues
  };
};
