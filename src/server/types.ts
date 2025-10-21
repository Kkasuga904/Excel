import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions';

export type CellValue = string | number | null | undefined;
export type CellMatrix = CellValue[][];

export interface CellData {
  address: string;
  values: CellMatrix;
}

export interface MessageHistoryItem {
  role: 'user' | 'assistant' | 'system';
  content: string;
}

export interface ExcelIntent {
  category:
    | 'statistics'
    | 'outlier_detection'
    | 'data_cleaning'
    | 'sorting'
    | 'reporting'
    | 'charting'
    | 'other';
  operations: string[];
  confidence: number;
  clarificationNeeded?: boolean;
}

export interface ExcelActionPlan {
  message: string;
  action: 'write' | 'analyze' | 'report' | 'chart' | 'none';
  intent: ExcelIntent;
  data?: {
    address?: string;
    values?: CellMatrix;
    chartType?: ChartType;
    metadata?: Record<string, unknown>;
  } | null;
  followUp?: string | null;
}

export interface ChatRequestBody {
  message: string;
  cellData?: CellData | null;
  history?: MessageHistoryItem[];
  mode?: 'basic' | 'advanced';
}

export interface StatisticsSummary {
  sum: string;
  average: string;
  max: string;
  min: string;
  median: string;
  count: number;
}

export type ChartType = 'LineChart' | 'BarChart' | 'PieChart';

export interface ChartData {
  labels: string[];
  datasets: Array<{ label: string; data: number[] }>;
  chartType: ChartType;
}

export interface PivotTableResult {
  columns: string[];
  rows: (string | number)[][];
}

export type ChatMessage = ChatCompletionMessageParam;
