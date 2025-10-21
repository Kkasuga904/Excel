import {
  prepareChartData,
  generatePivotTable,
  calculateMovingAverage,
  calculateCorrelation,
  performRegression,
  calculateZScores
} from '../src/server/excel-advanced';

describe('excel-advanced', () => {
  it('prepares chart data from matrix', () => {
    const matrix = [
      ['Month', 'Sales', 'Cost'],
      ['Jan', 100, 50],
      ['Feb', 120, 60]
    ];
    const chart = prepareChartData(matrix, 'BarChart');
    expect(chart).not.toBeNull();
    expect(chart?.labels).toEqual(['Sales', 'Cost']);
    expect(chart?.datasets[0].label).toBe('Jan');
    expect(chart?.chartType).toBe('BarChart');
  });

  it('builds pivot table aggregation', () => {
    const data = [
      ['East', 'A', 10],
      ['East', 'B', 20],
      ['West', 'A', 15]
    ];
    const pivot = generatePivotTable(data, 0, 1, 2);
    expect(pivot).not.toBeNull();
    expect(pivot?.rows.length).toBeGreaterThan(1);
  });

  it('computes moving average', () => {
    const series = [1, 2, 3, 4, 5];
    const moving = calculateMovingAverage(series, 3);
    expect(moving).toEqual([null, null, '2.00', '3.00', '4.00']);
  });

  it('calculates correlation', () => {
    const correlation = calculateCorrelation([1, 2, 3], [2, 4, 6]);
    expect(correlation).toBe('1.0000');
  });

  it('performs regression analysis', () => {
    const regression = performRegression([1, 2, 3], [2, 4, 6]);
    expect(regression).not.toBeNull();
    expect(regression?.slope).toBe('2.0000');
  });

  it('calculates z-scores and flags outliers', () => {
    const sample = [...Array(20).fill(1), 100];
    const scores = calculateZScores(sample);
    expect(scores).not.toBeNull();
    const flagged = scores?.find((score) => score.isOutlier);
    expect(flagged).toBeDefined();
  });
});
