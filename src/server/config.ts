import dotenv from 'dotenv';

dotenv.config();

const parseNumber = (value: string | undefined, fallback: number): number => {
  if (!value) {
    return fallback;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};

const parseList = (value: string | undefined): string[] =>
  value
    ?.split(',')
    .map((item) => item.trim())
    .filter((item) => item.length > 0) ?? [];

export const port = parseNumber(process.env.PORT, 3001);
export const nodeEnv = process.env.NODE_ENV ?? 'development';
export const corsOrigin = process.env.CORS_ORIGIN ?? '';

const allowedOrigins = parseList(corsOrigin);
const rateLimitWindowMs = parseNumber(process.env.RATE_LIMIT_WINDOW_MS, 15 * 60 * 1000);
const rateLimitMax = parseNumber(process.env.RATE_LIMIT_MAX, 100);
const requestTimeoutMs = parseNumber(process.env.REQUEST_TIMEOUT_MS, 60_000);
const maxTokens = parseNumber(process.env.MAX_TOKENS, 2000);
const temperature = parseNumber(process.env.OPENAI_TEMPERATURE, 0.7);

const config = {
  port,
  nodeEnv,
  corsOrigin,
  allowedOrigins,
  rateLimitWindowMs,
  rateLimitMax,
  requestTimeoutMs,
  openAIApiKey: process.env.OPENAI_API_KEY ?? '',
  openAIModel: process.env.OPENAI_MODEL ?? 'gpt-4',
  maxTokens,
  temperature
};

export type AppConfig = typeof config;

export default config;
