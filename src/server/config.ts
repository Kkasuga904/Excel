import 'dotenv/config';

export interface AppConfig {
  port: number;
  rateLimitWindowMs: number;
  rateLimitMax: number;
  requestTimeoutMs: number;
  allowedOrigins: string[];
  openAIApiKey: string;
  openAIModel: string;
  maxTokens: number;
  temperature: number;
}

const parseNumber = (value: string | undefined, fallback: number): number => {
  if (!value) {
    return fallback;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};

const parseOrigins = (value: string | undefined): string[] => {
  if (!value) {
    return [];
  }
  return value
    .split(',')
    .map((origin) => origin.trim())
    .filter(Boolean);
};

export const loadConfig = (): AppConfig => {
  const {
    PORT,
    RATE_LIMIT_WINDOW_MS,
    RATE_LIMIT_MAX,
    REQUEST_TIMEOUT_MS,
    CORS_ORIGIN,
    OPENAI_API_KEY,
    OPENAI_MODEL,
    MAX_TOKENS,
    OPENAI_TEMPERATURE
  } = process.env;

  return {
    port: parseNumber(PORT, 3001),
    rateLimitWindowMs: parseNumber(RATE_LIMIT_WINDOW_MS, 15 * 60 * 1000),
    rateLimitMax: parseNumber(RATE_LIMIT_MAX, 100),
    requestTimeoutMs: parseNumber(REQUEST_TIMEOUT_MS, 60000),
    allowedOrigins: parseOrigins(CORS_ORIGIN),
    openAIApiKey: OPENAI_API_KEY ?? '',
    openAIModel: OPENAI_MODEL ?? 'gpt-4',
    maxTokens: parseNumber(MAX_TOKENS, 2000),
    temperature: parseNumber(OPENAI_TEMPERATURE, 0.7)
  };
};
