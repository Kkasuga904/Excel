import express from 'express';
import helmet from 'helmet';
import config, { port, nodeEnv } from './config';
import {
  createRateLimiter,
  createCorsMiddleware,
  createTimeoutMiddleware,
  sanitizeRequestBody,
  validateChatInput,
  requireOpenAIKey,
  requestLogger,
  errorHandler,
  notFoundHandler
} from './middleware';
import { createChatService } from './openai';
import { createAdvancedChatService } from './openai-enhanced';
import {
  calculateStatistics,
  detectOutliers,
  generateMonthlyReport,
  fillBlankCells,
  sortData,
  detectDuplicates
} from './excel-helpers';
import { prepareChartData } from './excel-advanced';
import type { ChatRequestBody, CellData } from './types';

const app = express();
app.disable('x-powered-by');
app.set('trust proxy', 1);

app.use(express.json({ limit: '10mb' }));
app.use(helmet());
app.use(createCorsMiddleware(config));
app.use(requestLogger);
app.use(createTimeoutMiddleware(config.requestTimeoutMs));
app.use(sanitizeRequestBody);
app.use(createRateLimiter(config));

const chatService = createChatService(config);
const advancedChatService = createAdvancedChatService(config);

app.get('/health', (_req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    model: config.openAIModel
  });
});

app.post('/api/chat', validateChatInput, requireOpenAIKey(config), async (req, res, next) => {
  try {
    const body = req.body as ChatRequestBody;
    const response = await chatService.chat(body);
    res.json(response);
  } catch (error) {
    next(error);
  }
});

app.post('/api/chat/advanced', validateChatInput, requireOpenAIKey(config), async (req, res, next) => {
  try {
    const body = req.body as ChatRequestBody;
    const response = await advancedChatService.chat(body);
    res.json(response);
  } catch (error) {
    next(error);
  }
});

app.post('/api/analyze', requireOpenAIKey(config), async (req, res, next) => {
  try {
    const body = req.body as { cellData?: CellData };
    if (!body.cellData) {
      return res.status(400).json({ error: 'セルデータが必要です。' });
    }

    const { values, address } = body.cellData;
    const stats = calculateStatistics(values);
    const outliers = detectOutliers(values);
    const duplicates = detectDuplicates(values);
    const report = generateMonthlyReport(values, body.cellData);

    res.json({
      address,
      statistics: stats,
      outliers,
      duplicates,
      report
    });
  } catch (error) {
    next(error);
  }
});

app.post('/api/tools/fill-blanks', requireOpenAIKey(config), (req, res) => {
  const { values } = req.body as { values: CellData['values'] };
  res.json({ values: fillBlankCells(values) });
});

app.post('/api/tools/sort', requireOpenAIKey(config), (req, res) => {
  const { values, columnIndex, ascending } = req.body as {
    values: CellData['values'];
    columnIndex: number;
    ascending?: boolean;
  };

  res.json({
    values: sortData(values, columnIndex, ascending !== false)
  });
});

app.post('/api/tools/chart', requireOpenAIKey(config), (req, res) => {
  const { values, chartType } = req.body as {
    values: CellData['values'];
    chartType?: 'LineChart' | 'BarChart' | 'PieChart';
  };
  const chart = prepareChartData(values, chartType ?? 'LineChart');
  res.json({ chart });
});

app.use(notFoundHandler);
app.use(errorHandler);

if (nodeEnv !== "test") {
  app.listen(port, () => {
    console.log("========================================");
    console.log(" Excel AI Add-in Backend Server ");
    console.log("========================================");
    console.log(`Port: ${port}`);
    console.log(`Environment: ${nodeEnv}`);
    console.log(`OpenAI API Key: ${config.openAIApiKey ? 'configured' : 'missing'}`);
    console.log("========================================");
  });
}

export default app;


