import { OpenAI } from 'openai';
import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions';
import { AppConfig } from './config';
import { ChatRequestBody, ExcelActionPlan, MessageHistoryItem } from './types';
import { sanitizeCellData } from './excel-helpers';

const ADVANCED_PROMPT = [
  'You are an advanced analytics assistant for Microsoft Excel and must reply in Japanese.',
  'Provide detailed analytical reasoning, suggest step by step operations, and respect the JSON response schema defined for the basic assistant.',
  'Include tips for implementing the plan with Office Scripts when appropriate.'
].join('\n');

const toMessages = (history: MessageHistoryItem[] = []): ChatCompletionMessageParam[] =>
  history.map((entry) => ({ role: entry.role, content: entry.content }));

const formatUserMessage = (body: ChatRequestBody): string => {
  const parts = [`高度分析リクエスト: ${body.message}`];
  if (body.cellData) {
    const sanitized = sanitizeCellData(body.cellData);
    if (sanitized) {
      parts.push(
        `セル範囲: ${sanitized.address}`,
        `値サンプル: ${JSON.stringify(sanitized.values.slice(0, 15))}`
      );
    }
  }
  return parts.join('\n');
};

export class AdvancedChatService {
  constructor(private readonly client: OpenAI, private readonly config: AppConfig) {}

  async chat(body: ChatRequestBody): Promise<ExcelActionPlan> {
    const response = await this.client.chat.completions.create({
      model: this.config.openAIModel,
      max_tokens: this.config.maxTokens,
      temperature: this.config.temperature,
      messages: [
        { role: 'system', content: ADVANCED_PROMPT },
        ...toMessages(body.history),
        { role: 'user', content: formatUserMessage(body) }
      ]
    });

    const content = response.choices[0]?.message?.content;
    if (!content) {
      throw new Error('高度分析の応答が取得できませんでした。');
    }

    try {
      return JSON.parse(content) as ExcelActionPlan;
    } catch (error) {
      console.error('Failed to parse advanced AI response:', error);
      throw new Error('AIレスポンスの解析に失敗しました。');
    }
  }
}

export const createAdvancedChatService = (config: AppConfig): AdvancedChatService => {
  const client = new OpenAI({ apiKey: config.openAIApiKey });
  return new AdvancedChatService(client, config);
};
