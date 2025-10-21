/**
 * OpenAI API連携モジュール
 */

const { OpenAI } = require('openai');
const excelHelpers = require('./excel-helpers');

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

/**
 * システムプロンプト
 */
const SYSTEM_PROMPT = `あなたはExcel用のAIアシスタントです。
ユーザーの自然言語での指示を理解し、Excelでのデータ分析・操作・レポート作成を支援します。

利用可能な機能:
1. データ分析: 合計、平均、最大値、最小値、中央値の計算
2. 異常値検出: 外れ値の検出
3. データ操作: 空白セルの補完、重複削除、ソート
4. レポート生成: 月次レポートの作成
5. グラフ生成: 折れ線グラフ、円グラフの作成

返答形式（必ずJSON形式で返してください）:
{
  "message": "ユーザーへの説明（日本語）",
  "action": "write" | "analyze" | "report" | "chart" | "none",
  "data": {
    "address": "セルアドレス（例: A1:B10）",
    "values": [[...]]
  }
}

重要な注意事項:
- messageは必ず日本語で、ユーザーにわかりやすく説明してください
- actionは実行する操作を指定します
- dataはセルに書き込む結果を含めます
- エラーが発生した場合、messageにエラー内容を記述し、actionは"none"にしてください
- ユーザーの指示が曖昧な場合は、clarificationを求めてください`;

/**
 * AIチャットを実行
 */
async function chat(userMessage, cellData, messageHistory = []) {
  try {
    // メッセージ履歴を構築
    const messages = [
      ...messageHistory,
      {
        role: 'user',
        content: formatUserMessage(userMessage, cellData)
      }
    ];

    // OpenAI APIを呼び出し
    const response = await client.messages.create({
      model: process.env.OPENAI_MODEL || 'gpt-4',
      max_tokens: parseInt(process.env.MAX_TOKENS || '2000'),
      system: SYSTEM_PROMPT,
      messages: messages
    });

    const content = response.content[0];
    if (content.type !== 'text') {
      throw new Error('Unexpected response type from OpenAI');
    }

    // JSONレスポンスをパース
    const result = parseAIResponse(content.text);

    // セルデータに基づいて結果を処理
    if (result.action !== 'none' && cellData) {
      result.data = processAction(result.action, cellData);
    }

    return result;
  } catch (error) {
    console.error('OpenAI API error:', error);
    throw new Error(`AI処理エラー: ${error.message}`);
  }
}

/**
 * ユーザーメッセージをフォーマット
 */
function formatUserMessage(userMessage, cellData) {
  let formattedMessage = userMessage;

  if (cellData) {
    formattedMessage += `\n\n現在選択されているセルデータ:\nアドレス: ${cellData.address}\nデータ:\n${JSON.stringify(cellData.values)}`;
  }

  return formattedMessage;
}

/**
 * AIレスポンスをパース
 */
function parseAIResponse(responseText) {
  try {
    // JSONブロックを抽出
    const jsonMatch = responseText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return {
        message: responseText,
        action: 'none',
        data: null
      };
    }

    const parsed = JSON.parse(jsonMatch[0]);
    return {
      message: parsed.message || responseText,
      action: parsed.action || 'none',
      data: parsed.data || null
    };
  } catch (error) {
    console.error('Failed to parse AI response:', error);
    return {
      message: responseText,
      action: 'none',
      data: null
    };
  }
}

/**
 * アクションを処理
 */
function processAction(action, cellData) {
  try {
    switch (action) {
      case 'analyze':
        return processAnalyze(cellData);
      case 'report':
        return processReport(cellData);
      case 'chart':
        return processChart(cellData);
      default:
        return null;
    }
  } catch (error) {
    console.error('Error processing action:', error);
    return null;
  }
}

/**
 * 分析アクションを処理
 */
function processAnalyze(cellData) {
  const stats = excelHelpers.calculateStatistics(cellData.values);
  if (!stats) {
    return null;
  }

  const resultAddress = getNextRowAddress(cellData.address);
  const resultValues = [
    ['分析結果'],
    ['合計', stats.sum],
    ['平均', stats.average],
    ['最大値', stats.max],
    ['最小値', stats.min],
    ['中央値', stats.median],
    ['データ数', stats.count]
  ];

  return {
    address: resultAddress,
    values: resultValues
  };
}

/**
 * レポートアクションを処理
 */
function processReport(cellData) {
  const report = excelHelpers.generateMonthlyReport(cellData.values, cellData.address);
  if (!report) {
    return null;
  }

  return {
    address: 'A1',
    values: report
  };
}

/**
 * グラフアクションを処理
 */
function processChart(cellData) {
  // グラフ生成は別途Office.jsで処理するため、ここではデータのみ返す
  return {
    address: cellData.address,
    values: cellData.values,
    chartType: 'LineChart'
  };
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
  chat
};

