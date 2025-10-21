/**
 * エラーハンドリングとセキュリティミドルウェア
 */

const rateLimit = require('express-rate-limit');

/**
 * レート制限ミドルウェア
 */
const createRateLimiter = (windowMs = 15 * 60 * 1000, max = 100) => {
  return rateLimit({
    windowMs: windowMs,
    max: max,
    message: '利用制限に達しました。しばらく待ってから再度お試しください。',
    standardHeaders: true,
    legacyHeaders: false,
    skip: (req) => {
      // ヘルスチェックはレート制限をスキップ
      return req.path === '/health';
    }
  });
};

/**
 * 入力検証ミドルウェア
 */
const validateInput = (req, res, next) => {
  const { message, cellData } = req.body;

  // メッセージの検証
  if (!message || typeof message !== 'string') {
    return res.status(400).json({
      error: 'メッセージが必要です'
    });
  }

  if (message.length > 1000) {
    return res.status(400).json({
      error: 'メッセージは1000文字以内にしてください'
    });
  }

  if (message.trim().length === 0) {
    return res.status(400).json({
      error: 'メッセージが空です'
    });
  }

  // セルデータの検証
  if (cellData) {
    if (!cellData.address || typeof cellData.address !== 'string') {
      return res.status(400).json({
        error: 'セルアドレスが無効です'
      });
    }

    if (!Array.isArray(cellData.values)) {
      return res.status(400).json({
        error: 'セルデータが無効です'
      });
    }

    // セルデータのサイズをチェック
    const dataSize = JSON.stringify(cellData).length;
    if (dataSize > 10 * 1024 * 1024) { // 10MB
      return res.status(400).json({
        error: 'セルデータが大きすぎます（最大10MB）'
      });
    }
  }

  next();
};

/**
 * セキュリティヘッダーミドルウェア
 */
const securityHeaders = (req, res, next) => {
  // XSS対策
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  res.setHeader('X-XSS-Protection', '1; mode=block');

  // CSP設定
  res.setHeader(
    'Content-Security-Policy',
    "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'"
  );

  // HSTS設定（本番環境のみ）
  if (process.env.NODE_ENV === 'production') {
    res.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains');
  }

  next();
};

/**
 * ログミドルウェア
 */
const requestLogger = (req, res, next) => {
  const startTime = Date.now();

  // レスポンス終了時にログ出力
  res.on('finish', () => {
    const duration = Date.now() - startTime;
    const logLevel = res.statusCode >= 400 ? 'error' : 'info';

    console.log(`[${logLevel.toUpperCase()}] ${req.method} ${req.path} - ${res.statusCode} (${duration}ms)`);

    // エラーの詳細ログ
    if (res.statusCode >= 400) {
      console.error(`Request body: ${JSON.stringify(req.body)}`);
    }
  });

  next();
};

/**
 * エラーハンドリングミドルウェア
 */
const errorHandler = (err, req, res, next) => {
  console.error('Error:', err);

  // デフォルトエラーレスポンス
  let statusCode = err.statusCode || 500;
  let message = err.message || 'サーバーエラーが発生しました';

  // 既知のエラータイプ
  if (err.name === 'ValidationError') {
    statusCode = 400;
    message = 'バリデーションエラーが発生しました';
  } else if (err.name === 'UnauthorizedError') {
    statusCode = 401;
    message = '認証に失敗しました';
  } else if (err.name === 'ForbiddenError') {
    statusCode = 403;
    message = 'アクセスが拒否されました';
  } else if (err.name === 'NotFoundError') {
    statusCode = 404;
    message = 'リソースが見つかりません';
  } else if (err.name === 'RateLimitError') {
    statusCode = 429;
    message = '利用制限に達しました。しばらく待ってから再度お試しください';
  }

  // 本番環境ではエラー詳細を隠す
  const responseMessage = process.env.NODE_ENV === 'production'
    ? message
    : `${message}: ${err.message}`;

  res.status(statusCode).json({
    error: responseMessage,
    ...(process.env.NODE_ENV !== 'production' && { details: err.stack })
  });
};

/**
 * 404ハンドラ
 */
const notFoundHandler = (req, res) => {
  res.status(404).json({
    error: 'エンドポイントが見つかりません'
  });
};

/**
 * APIキー検証ミドルウェア
 */
const validateAPIKey = (req, res, next) => {
  // 環境変数でAPIキーが設定されているか確認
  if (!process.env.OPENAI_API_KEY) {
    return res.status(500).json({
      error: 'OpenAI APIキーが設定されていません。README を確認してください。'
    });
  }

  next();
};

/**
 * リクエストタイムアウトミドルウェア
 */
const requestTimeout = (timeout = 30000) => {
  return (req, res, next) => {
    const timeoutId = setTimeout(() => {
      if (!res.headersSent) {
        res.status(408).json({
          error: 'リクエストがタイムアウトしました'
        });
      }
    }, timeout);

    res.on('finish', () => {
      clearTimeout(timeoutId);
    });

    next();
  };
};

/**
 * リクエストサニタイズミドルウェア
 */
const sanitizeInput = (req, res, next) => {
  if (req.body && req.body.message) {
    // HTMLタグを削除
    req.body.message = req.body.message.replace(/<[^>]*>/g, '');

    // 危険な文字を削除
    req.body.message = req.body.message.replace(/[<>\"']/g, '');
  }

  next();
};

/**
 * CORS設定
 */
const getCORSOptions = () => {
  const allowedOrigins = [
    'http://localhost:3000',
    'http://localhost:3001',
    'https://localhost:3000',
    'https://localhost:3001'
  ];

  if (process.env.CORS_ORIGIN) {
    allowedOrigins.push(process.env.CORS_ORIGIN);
  }

  return {
    origin: (origin, callback) => {
      if (!origin || allowedOrigins.includes(origin)) {
        callback(null, true);
      } else {
        callback(new Error('CORSポリシーに違反しています'));
      }
    },
    credentials: true,
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    maxAge: 86400 // 24時間
  };
};

module.exports = {
  createRateLimiter,
  validateInput,
  securityHeaders,
  requestLogger,
  errorHandler,
  notFoundHandler,
  validateAPIKey,
  requestTimeout,
  sanitizeInput,
  getCORSOptions
};

