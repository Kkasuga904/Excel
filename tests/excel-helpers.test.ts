import {
  calculateStatistics,
  detectOutliers,
  fillBlankCells,
  generateMonthlyReport,
  sortData
} from '../src/server/excel-helpers';

describe('excel-helpers', () => {
  const values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ];

  it('calculates statistics with sorted median', () => {
    const stats = calculateStatistics(values);
    expect(stats).not.toBeNull();
    expect(stats?.median).toBe('5.00');
    expect(stats?.sum).toBe('45.00');
    expect(stats?.count).toBe(9);
  });

  it('detects outliers using IQR', () => {
    const sample = [
      [10],
      [12],
      [14],
      [15],
      [99]
    ];
    const outliers = detectOutliers(sample);
    expect(outliers).toContain(99);
    expect(outliers.length).toBe(1);
  });

  it('fills blank cells by interpolation', () => {
    const input = [
      [1, null, 3],
      [4, null, 6],
      [7, 8, 9]
    ];
    const result = fillBlankCells(input);
    expect(result[0][1]).toBe('8.00');
    expect(result[1][1]).toBe('8.00');
  });

  it('generates monthly report when statistics available', () => {
    const report = generateMonthlyReport(values, { address: 'A1:C3', values });
    expect(report).not.toBeNull();
    expect(report?.[0][0]).toBe('月次レポート');
    expect(report?.[5][1]).toBe('45.00');
  });

  it('sorts data by numeric column', () => {
    const sorted = sortData(
      [
        ['A', 3],
        ['B', 1],
        ['C', 2]
      ],
      1
    );
    expect(sorted[0][0]).toBe('B');
    expect(sorted[2][0]).toBe('A');
  });
});
