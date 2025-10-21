/**
 * Express サーバー
 * Excel AI アドインのバックエンド
 */

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { chat } = require('./openai');
const {
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
} = require('./middleware');

const app = express();
const PORT = process.env.PORT || 3001;

// ミドルウェア
app.use(express.json({ limit: '10mb' }));
app.use(cors(getCORSOptions()));
app.use(securityHeaders);
app.use(requestLogger);
app.use(requestTimeout(30000));
app.use(sanitizeInput);
app.use(createRateLimiter(15 * 60 * 1000, 100));

/**
 * ヘルスチェックエンドポイント
 */
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

/**
 * チャットエンドポイント
 */
app.post('/api/chat', validateInput, validateAPIKey, async (req, res) => {
  try {
    const { message, cellData, messageHistory } = req.body;

    // AIチャットを実行
    const result = await chat(message, cellData, messageHistory || []);

    res.json({
      message: result.message,
      action: result.action,
      data: result.data
    });
  } catch (error) {
    console.error('Chat API error:', error);
    res.status(500).json({
      error: error.message || 'チャット処理中にエラーが発生しました'
    });
  }
});

/**
 * 分析エンドポイント
 */
app.post('/api/analyze', async (req, res) => {
  try {
    const { cellData } = req.body;

    if (!cellData || !cellData.values) {
      return res.status(400).json({
        error: 'セルデータが必要です'
      });
    }

    const excelHelpers = require('./excel-helpers');
    const stats = excelHelpers.calculateStatistics(cellData.values);
    const outliers = excelHelpers.detectOutliers(cellData.values);

    res.json({
      statistics: stats,
      outliers: outliers,
      address: cellData.address
    });
  } catch (error) {
    console.error('Analyze API error:', error);
    res.status(500).json({
      error: error.message || '分析処理中にエラーが発生しました'
    });
  }
});

/**
 * エラーハンドリング
 */
app.use(notFoundHandler);
app.use(errorHandler);

// サーバー起動
app.listen(PORT, () => {
  console.log(`\n========================================`);
  console.log(`Excel AI アドイン バックエンドサーバー`);
  console.log(`========================================`);
  console.log(`ポート: ${PORT}`);
  console.log(`環境: ${process.env.NODE_ENV || 'development'}`);
  console.log(`OpenAI API: ${process.env.OPENAI_API_KEY ? '設定済み' : '未設定'}`);
  console.log(`========================================\n`);
});

// グレースフルシャットダウン
process.on('SIGTERM', () => {
  console.log('SIGTERM received, shutting down gracefully...');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT received, shutting down gracefully...');
  process.exit(0);
});

