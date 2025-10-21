import { OpenAI } from 'openai';
import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions';
import { AppConfig } from './config';
import {
  ChatRequestBody,
  ExcelActionPlan,
  ExcelIntent,
  MessageHistoryItem
} from './types';
import { sanitizeCellData } from './excel-helpers';
import { prepareChartData } from './excel-advanced';

const SYSTEM_PROMPT = [
  'You are a Microsoft Excel assistant that must reply in natural Japanese.',
  'Analyse the user request, determine the appropriate Excel intent, and respond with JSON only.',
  'Categorise the intent into one of the following domains:',
  '- statistics: total, average, median, min/max, basic descriptive analytics',
  '- outlier_detection: IQR, z-score, anomaly checks',
  '- data_cleaning: fill blanks, deduplicate, normalise',
  '- sorting: ordering, filtering, re-arranging rows',
  '- reporting: summaries, monthly reports, formatted output',
  '- charting: line, bar, pie charts and related data shaping',
  '- other: anything that does not match the above.',
  '',
  'Return JSON with the exact structure:',
  '{',
  '  "message": "Japanese explanation for the user",',
  '  "action": "write" | "analyze" | "report" | "chart" | "none",',
  '  "intent": {',
  '    "category": "statistics" | "outlier_detection" | "data_cleaning" | "sorting" | "reporting" | "charting" | "other",',
  '    "operations": ["operation name", ...],',
  '    "confidence": number between 0 and 1,',
  '    "clarificationNeeded": boolean',
  '  },',
  '  "data": {',
  '    "address": "target range such as A1:B10",',
  '    "values": [[...]],',
  '    "chartType": "LineChart" | "BarChart" | "PieChart" | null,',
  '    "metadata": { "any": "additional details" }',
  '  },',
  '  "followUp": null or a clarification question in Japanese',
  '}',
  '',
  'Only emit JSON. Do not include code fences or explanatory text outside the JSON structure.',
  'Set clarificationNeeded to true and provide a follow-up question when the request is ambiguous or incomplete.',
  'Avoid suggesting unsafe or irreversible actions.',
  'When generating chart plans include chartType.',
  'Keep responses polite and concise.'
].join('\n');

const extractJson = (content: string): string => {
  const match = content.match(/\{[\s\S]*\}/);
  return match ? match[0] : content;
};

const toHistoryMessages = (history: MessageHistoryItem[] = []): ChatCompletionMessageParam[] =>
  history.map((entry) => ({
    role: entry.role,
    content: entry.content
  }));

const truncateValues = (values: unknown[][], limit = 10): unknown[][] => values.slice(0, limit);

const formatUserMessage = (body: ChatRequestBody): string => {
  const lines = [`謖・､ｺ: ${body.message}`];
  if (body.cellData) {
    const sanitized = sanitizeCellData(body.cellData);
    if (sanitized) {
      lines.push(
        `繧ｻ繝ｫ遽・峇: ${sanitized.address}`,
        `蛟､繝励Ξ繝薙Η繝ｼ: ${JSON.stringify(truncateValues(sanitized.values))}`
      );
    }
  }
  return lines.join('\n');
};

const fallbackPlan = (message: string): ExcelActionPlan => ({
  message,
  action: 'none',
  intent: {
    category: 'other',
    operations: [],
    confidence: 0,
    clarificationNeeded: true
  },
  data: null,
  followUp: 'Please provide more detail about your request.'
});
const normaliseIntent = (intent?: Partial<ExcelIntent>): ExcelIntent => {
  const rawOperations = intent?.operations;
  const operations = Array.isArray(rawOperations)
    ? rawOperations.map((operation) => String(operation))
    : [];

  const confidenceValue =
    typeof intent?.confidence === 'number' && intent.confidence >= 0 && intent.confidence <= 1
      ? intent.confidence
      : 0.5;

  return {
    category: intent?.category ?? 'other',
    operations,
    confidence: confidenceValue,
    clarificationNeeded: Boolean(intent?.clarificationNeeded)
  };
};

export class OpenAIChatService {
  constructor(private readonly client: OpenAI, private readonly config: AppConfig) {}

  async chat(body: ChatRequestBody): Promise<ExcelActionPlan> {
    const messages: ChatCompletionMessageParam[] = [
      { role: 'system', content: SYSTEM_PROMPT },
      ...toHistoryMessages(body.history),
      { role: 'user', content: formatUserMessage(body) }
    ];

    const response = await this.client.chat.completions.create({
      model: this.config.openAIModel,
      max_tokens: this.config.maxTokens,
      messages,
      temperature: this.config.temperature
    });

    const content = response.choices[0]?.message?.content;
    if (!content) {
      throw new Error('Failed to receive a response from OpenAI.');
    }

    const raw = extractJson(content);

    try {
      const parsed = JSON.parse(raw) as Partial<ExcelActionPlan>;
      const intent = normaliseIntent(parsed.intent);
      const action =
        parsed.action === 'write' ||
        parsed.action === 'analyze' ||
        parsed.action === 'report' ||
        parsed.action === 'chart'
          ? parsed.action
          : 'none';

      const plan: ExcelActionPlan = {
        message: typeof parsed.message === 'string' ? parsed.message : 'Result generated.',
        action,
        intent,
        data: parsed.data ?? null,
        followUp: 'Please provide more detail about your request.'
      };

      if (plan.action === 'chart' && !plan.data?.values && body.cellData) {
        const chartData = prepareChartData(body.cellData.values, plan.data?.chartType ?? 'LineChart');
        if (chartData) {
          plan.data = {
            ...plan.data,
            values: [
              ['繝ｩ繝吶Ν', ...chartData.labels],
              ...chartData.datasets.map((dataset) => [dataset.label, ...dataset.data])
            ]
          };
        }
      }

      return plan;
    } catch (error) {
      console.error('Failed to parse AI response:', error);
      return fallbackPlan('Failed to parse the AI response.');
    }
  }
}

export const createChatService = (config: AppConfig): OpenAIChatService => {
  const client = new OpenAI({ apiKey: config.openAIApiKey });
  return new OpenAIChatService(client, config);
};







