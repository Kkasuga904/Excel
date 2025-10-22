import cors, { CorsOptions } from 'cors';
import { Request, Response, NextFunction, RequestHandler } from 'express';
import rateLimit from 'express-rate-limit';
import sanitizeHtml from 'sanitize-html';
import { AppConfig } from './config';
import { ChatRequestBody } from './types';

export const createRateLimiter = (config: AppConfig): RequestHandler =>
  rateLimit({
    windowMs: config.rateLimitWindowMs,
    max: config.rateLimitMax,
    standardHeaders: true,
    legacyHeaders: false,
    skip: (req) => req.path === '/health'
  });

export const createCorsMiddleware = (config: AppConfig): RequestHandler => {
  const defaultOrigins = [
    'http://localhost:3000',
    'http://localhost:3001',
    'https://localhost:3000',
    'https://localhost:3001',
    'http://127.0.0.1:3000',
    'http://127.0.0.1:3001',
    'https://127.0.0.1:3000',
    'https://127.0.0.1:3001'
  ];

  const allowedOrigins = [...defaultOrigins, ...config.allowedOrigins];

  const corsOptions: CorsOptions = {
    origin(origin: string | undefined, callback: (err: Error | null, allow?: boolean) => void) {
      if (!origin || allowedOrigins.includes(origin)) {
        callback(null, true);
      } else {
        callback(new Error('CORSポリシーによりリクエストが拒否されました。'));
      }
    },
    credentials: true,
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    maxAge: 60 * 60 * 24
  };

  return cors(corsOptions);
};

export const createTimeoutMiddleware =
  (timeoutMs: number): RequestHandler =>
  (req, res, next) => {
    const timer = setTimeout(() => {
      if (!res.headersSent) {
        res.status(408).json({ error: 'リクエストがタイムアウトしました。' });
      }
    }, timeoutMs);

    res.on('finish', () => clearTimeout(timer));
    next();
  };

const sanitizeMessage = (message: string): string =>
  sanitizeHtml(message, {
    allowedTags: [],
    allowedAttributes: {}
  }).trim();

export const sanitizeRequestBody: RequestHandler = (req, _res, next) => {
  const body = req.body as Partial<ChatRequestBody> | undefined;
  if (body && typeof body.message === 'string') {
    body.message = sanitizeMessage(body.message);
  }
  if (body?.history) {
    body.history = body.history
      .filter((entry) => entry && entry.role && entry.content)
      .map((entry) => ({
        role: entry.role,
        content: sanitizeMessage(entry.content)
      }));
  }
  next();
};

export const validateChatInput: RequestHandler = (req, res, next) => {
  const body = req.body as ChatRequestBody;

  if (!body || typeof body.message !== 'string' || body.message.trim().length === 0) {
    return res.status(400).json({ error: 'メッセージは必須です。' });
  }

  if (body.message.length > 2000) {
    return res.status(400).json({ error: 'メッセージは2000文字以内にしてください。' });
  }

  if (body.cellData) {
    if (typeof body.cellData.address !== 'string' || !Array.isArray(body.cellData.values)) {
      return res.status(400).json({ error: 'セルデータの形式が正しくありません。' });
    }
    const serialized = JSON.stringify(body.cellData);
    if (serialized.length > 10 * 1024 * 1024) {
      return res.status(400).json({ error: 'セルデータが大きすぎます（最大10MB）。' });
    }
  }

  next();
};

export const requireOpenAIKey =
  (config: AppConfig): RequestHandler =>
  (_req, res, next) => {
    if (!config.openAIApiKey) {
      return res.status(500).json({
        error: 'OpenAI APIキーが設定されていません。環境変数 OPENAI_API_KEY を設定してください。'
      });
    }
    next();
  };

export const requestLogger: RequestHandler = (req, res, next) => {
  const startedAt = Date.now();
  res.on('finish', () => {
    const duration = Date.now() - startedAt;
    const status = res.statusCode;
    const level = status >= 500 ? 'ERROR' : status >= 400 ? 'WARN' : 'INFO';
    console.log(`[${level}] ${req.method} ${req.originalUrl} - ${status} (${duration}ms)`);
  });
  next();
};

export const errorHandler = (
  err: Error & { statusCode?: number },
  _req: Request,
  res: Response,
  _next: NextFunction
) => {
  console.error('Unhandled error:', err);
  const statusCode = err.statusCode ?? 500;
  const response = {
    error: statusCode >= 500 ? 'サーバー内部でエラーが発生しました。' : err.message
  };
  res.status(statusCode).json(response);
};

export const notFoundHandler: RequestHandler = (_req, res) => {
  res.status(404).json({ error: 'エンドポイントが存在しません。' });
};
