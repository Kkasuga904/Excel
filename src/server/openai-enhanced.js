/**
 * 拡張OpenAI API連携モジュール
 * より高度なプロンプトエンジニアリングと機能を提供
 */

const { OpenAI } = require('openai');
const excelHelpers = require('./excel-helpers');
const excelAdvanced = require('./excel-advanced');

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

/**
 * 高度なシステムプロンプト
 */
const ADVANCED_SYSTEM_PROMPT = `あなたはExcel用の高度なAIアシスタントです。
ユーザーの自然言語での指示を理解し、Excelでのデータ分析・操作・レポート作成を支援します。

【利用可能な機能】

1. データ分析
   - 統計情報（合計、平均、最大値、最小値、中央値）
   - 異常値検出（IQR法）
   - Z-スコア計算
   - 相関係数計算
   - 回帰分析

2. データ操作
   - 空白セルの補完（線形補間）
   - 重複行の削除
   - データのソート
   - データの正規化

3. レポート生成
   - 月次レポート
   - サマリーシート
   - ピボットテーブル

4. グラフ作成
   - 折れ線グラフ
   - 円グラフ
   - 棒グラフ
   - 散布図

5. 高度な分析
   - 移動平均計算
   - トレンド分析
   - 条件付き書式

【返答形式】

必ず以下のJSON形式で返してください：

{
  "message": "ユーザーへの説明（日本語、詳細で分かりやすく）",
  "action": "analyze" | "fill" | "sort" | "report" | "chart" | "pivot" | "regression" | "correlation" | "moving_average" | "none",
  "data": {
    "address": "セルアドレス（例: A1:B10）",
    "values": [[...]]
  },
  "metadata": {
    "description": "実行内容の説明",
    "parameters": {...}
  }
}

【重要な注意事項】

- messageは必ず日本語で、ユーザーにわかりやすく説明してください
- 数値は適切に丸めて表示してください
- エラーが発生した場合、messageにエラー内容を記述し、actionは"none"にしてください
- ユーザーの指示が曖昧な場合は、clarificationを求めてください
- 複雑な操作は段階的に説明してください
- セルアドレスは常に有効な形式で指定してください（例: A1:C10）`;

/**
 * 会話コンテキストを管理
 */
class ConversationManager {
  constructor(maxHistory = 10) {
    this.history = [];
    this.maxHistory = maxHistory;
  }

  addMessage(role, content) {
    this.history.push({ role, content });
    if (this.history.length > this.maxHistory) {
      this.history.shift();
    }
  }

  getHistory() {
    return this.history;
  }

  clear() {
    this.history = [];
  }
}

/**
 * 高度なAIチャットを実行
 */
async function advancedChat(userMessage, cellData, conversationManager) {
  try {
    // メッセージ履歴を構築
    const messages = [
      ...conversationManager.getHistory(),
      {
        role: 'user',
        content: formatAdvancedUserMessage(userMessage, cellData)
      }
    ];

    // OpenAI APIを呼び出し
    const response = await client.chat.completions.create({
      model: process.env.OPENAI_MODEL || 'gpt-4',
      max_tokens: parseInt(process.env.MAX_TOKENS || '2000'),
      messages: [
        { role: 'system', content: ADVANCED_SYSTEM_PROMPT },
        ...messages
      ],
      temperature: 0.7
    });

    const content = response.choices[0]?.message?.content;
    if (!content) {
      throw new Error('No response from OpenAI');
    }

    // JSONレスポンスをパース
    const result = parseAdvancedAIResponse(content);

    // 会話履歴に追加
    conversationManager.addMessage('user', userMessage);
    conversationManager.addMessage('assistant', result.message);

    // セルデータに基づいて結果を処理
    if (result.action !== 'none' && cellData) {
      result.data = processAdvancedAction(result.action, cellData, result.metadata);
    }

    return result;
  } catch (error) {
    console.error('Advanced OpenAI API error:', error);
    throw new Error(`AI処理エラー: ${error.message}`);
  }
}

/**
 * ユーザーメッセージをフォーマット（高度版）
 */
function formatAdvancedUserMessage(userMessage, cellData) {
  let formattedMessage = userMessage;

  if (cellData) {
    formattedMessage += `\n\n【現在選択されているセルデータ】\nアドレス: ${cellData.address}\nデータ:\n${JSON.stringify(cellData.values, null, 2)}`;

    // データの統計情報を追加
    const stats = excelHelpers.calculateStatistics(cellData.values);
    if (stats) {
      formattedMessage += `\n\n【統計情報】\n合計: ${stats.sum}\n平均: ${stats.average}\n最大値: ${stats.max}\n最小値: ${stats.min}\nデータ数: ${stats.count}`;
    }
  }

  return formattedMessage;
}

/**
 * AIレスポンスをパース（高度版）
 */
function parseAdvancedAIResponse(responseText) {
  try {
    // JSONブロックを抽出
    const jsonMatch = responseText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return {
        message: responseText,
        action: 'none',
        data: null,
        metadata: {}
      };
    }

    const parsed = JSON.parse(jsonMatch[0]);
    return {
      message: parsed.message || responseText,
      action: parsed.action || 'none',
      data: parsed.data || null,
      metadata: parsed.metadata || {}
    };
  } catch (error) {
    console.error('Failed to parse advanced AI response:', error);
    return {
      message: responseText,
      action: 'none',
      data: null,
      metadata: {}
    };
  }
}

/**
 * 高度なアクションを処理
 */
function processAdvancedAction(action, cellData, metadata = {}) {
  try {
    switch (action) {
      case 'analyze':
        return processAnalyzeAdvanced(cellData);
      case 'fill':
        return processFill(cellData);
      case 'sort':
        return processSort(cellData, metadata);
      case 'report':
        return processReportAdvanced(cellData);
      case 'chart':
        return processChartAdvanced(cellData, metadata);
      case 'pivot':
        return processPivot(cellData, metadata);
      case 'regression':
        return processRegression(cellData, metadata);
      case 'correlation':
        return processCorrelation(cellData, metadata);
      case 'moving_average':
        return processMovingAverage(cellData, metadata);
      default:
        return null;
    }
  } catch (error) {
    console.error('Error processing advanced action:', error);
    return null;
  }
}

/**
 * 分析アクション（高度版）
 */
function processAnalyzeAdvanced(cellData) {
  const stats = excelHelpers.calculateStatistics(cellData.values);
  const outliers = excelHelpers.detectOutliers(cellData.values);
  const zScores = excelAdvanced.calculateZScores(
    cellData.values.flat().map((v) => parseFloat(v)).filter((v) => !isNaN(v))
  );

  if (!stats) {
    return null;
  }

  const resultAddress = getNextRowAddress(cellData.address);
  const resultValues = [
    ['分析結果'],
    [''],
    ['基本統計情報'],
    ['合計', stats.sum],
    ['平均', stats.average],
    ['最大値', stats.max],
    ['最小値', stats.min],
    ['中央値', stats.median],
    ['データ数', stats.count],
    [''],
    ['異常値検出'],
    ['異常値の数', outliers.length],
    ['異常値', outliers.join(', ') || 'なし']
  ];

  return {
    address: resultAddress,
    values: resultValues
  };
}

/**
 * 空白補完アクション
 */
function processFill(cellData) {
  const filled = excelHelpers.fillBlankCells(cellData.values);
  return {
    address: cellData.address,
    values: filled
  };
}

/**
 * ソートアクション
 */
function processSort(cellData, metadata) {
  const columnIndex = metadata.columnIndex || 0;
  const ascending = metadata.ascending !== false;
  const sorted = excelHelpers.sortData(cellData.values, columnIndex, ascending);

  return {
    address: cellData.address,
    values: sorted
  };
}

/**
 * レポート生成アクション（高度版）
 */
function processReportAdvanced(cellData) {
  const stats = excelHelpers.calculateStatistics(cellData.values);
  const report = excelHelpers.generateMonthlyReport(cellData.values, cellData.address);

  return {
    address: 'A1',
    values: report
  };
}

/**
 * グラフ生成アクション（高度版）
 */
function processChartAdvanced(cellData, metadata) {
  const chartType = metadata.chartType || 'line';
  const chartData = excelAdvanced.prepareChartData(cellData.values, chartType);

  return {
    address: cellData.address,
    values: cellData.values,
    chartData: chartData,
    chartType: chartType
  };
}

/**
 * ピボットテーブル処理
 */
function processPivot(cellData, metadata) {
  const rowIndex = metadata.rowIndex || 0;
  const colIndex = metadata.colIndex || 1;
  const valueIndex = metadata.valueIndex || 2;

  const pivot = excelAdvanced.generatePivotTable(
    cellData.values,
    rowIndex,
    colIndex,
    valueIndex
  );

  return {
    address: 'A1',
    values: pivot
  };
}

/**
 * 回帰分析処理
 */
function processRegression(cellData, metadata) {
  const xColIndex = metadata.xColIndex || 0;
  const yColIndex = metadata.yColIndex || 1;

  const xValues = cellData.values.map((row) => parseFloat(row[xColIndex])).filter((v) => !isNaN(v));
  const yValues = cellData.values.map((row) => parseFloat(row[yColIndex])).filter((v) => !isNaN(v));

  const regression = excelAdvanced.performRegression(xValues, yValues);

  if (!regression) {
    return null;
  }

  const resultAddress = getNextRowAddress(cellData.address);
  const resultValues = [
    ['回帰分析結果'],
    [''],
    ['方程式', regression.equation],
    ['傾き', regression.slope],
    ['切片', regression.intercept],
    ['R二乗', regression.rSquared]
  ];

  return {
    address: resultAddress,
    values: resultValues
  };
}

/**
 * 相関係数処理
 */
function processCorrelation(cellData, metadata) {
  const col1Index = metadata.col1Index || 0;
  const col2Index = metadata.col2Index || 1;

  const series1 = cellData.values.map((row) => parseFloat(row[col1Index])).filter((v) => !isNaN(v));
  const series2 = cellData.values.map((row) => parseFloat(row[col2Index])).filter((v) => !isNaN(v));

  const correlation = excelAdvanced.calculateCorrelation(series1, series2);

  if (correlation === null) {
    return null;
  }

  const resultAddress = getNextRowAddress(cellData.address);
  const resultValues = [
    ['相関係数分析'],
    [''],
    ['相関係数', correlation],
    ['解釈', interpretCorrelation(parseFloat(correlation))]
  ];

  return {
    address: resultAddress,
    values: resultValues
  };
}

/**
 * 移動平均処理
 */
function processMovingAverage(cellData, metadata) {
  const period = metadata.period || 3;
  const colIndex = metadata.colIndex || 0;

  const values = cellData.values.map((row) => row[colIndex]);
  const movingAvg = excelAdvanced.calculateMovingAverage(values, period);

  if (!movingAvg) {
    return null;
  }

  const resultAddress = getNextRowAddress(cellData.address);
  const resultValues = movingAvg.map((v) => [v]);

  return {
    address: resultAddress,
    values: resultValues
  };
}

/**
 * 相関係数を解釈
 */
function interpretCorrelation(correlation) {
  const absCorr = Math.abs(correlation);
  if (absCorr >= 0.8) {
    return '非常に強い相関';
  } else if (absCorr >= 0.6) {
    return '強い相関';
  } else if (absCorr >= 0.4) {
    return '中程度の相関';
  } else if (absCorr >= 0.2) {
    return '弱い相関';
  } else {
    return 'ほぼ相関なし';
  }
}

/**
 * 次の行のアドレスを取得
 */
function getNextRowAddress(address) {
  const match = address.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!match) {
    return 'A1';
  }

  const endRow = parseInt(match[4]);
  const nextRow = endRow + 2;
  return `A${nextRow}`;
}

module.exports = {
  advancedChat,
  ConversationManager,
  formatAdvancedUserMessage,
  parseAdvancedAIResponse,
  processAdvancedAction
};

