# アーキテクチャドキュメント

このドキュメントでは、Excel AI チャットアシスタントの技術的なアーキテクチャを詳細に説明します。

## システム概要

Excel AI チャットアシスタントは、Microsoft Excel用のAI駆動型アドインです。ユーザーが自然言語でデータ分析や操作を指示すると、AIが適切な処理を実行し、結果をExcelに自動出力します。

### アーキテクチャ図

```
┌─────────────────────────────────────────────────────────────┐
│                    Microsoft Excel                           │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  Excel AI チャットアドイン（Office.js）              │   │
│  │  ┌──────────────────────────────────────────────┐   │   │
│  │  │  React UI（チャットインターフェース）        │   │   │
│  │  │  - ChatMessage コンポーネント                │   │   │
│  │  │  - ChatInput コンポーネント                  │   │   │
│  │  │  - LoadingSpinner コンポーネント            │   │   │
│  │  └──────────────────────────────────────────────┘   │   │
│  │  ┌──────────────────────────────────────────────┐   │   │
│  │  │  Office.js 連携層                           │   │   │
│  │  │  - getSelectedData()                         │   │   │
│  │  │  - writeToCell()                            │   │   │
│  │  └──────────────────────────────────────────────┘   │   │
│  └──────────────────────────────────────────────────────┘   │
└────────────┬─────────────────────────────────────────────────┘
             │ HTTP/HTTPS
             ↓
┌─────────────────────────────────────────────────────────────┐
│              バックエンドサーバー（Express）                 │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  API エンドポイント                                  │   │
│  │  - POST /api/chat                                   │   │
│  │  - POST /api/analyze                               │   │
│  │  - GET /health                                     │   │
│  └──────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  ミドルウェア層                                      │   │
│  │  - 認証・認可                                       │   │
│  │  - レート制限                                       │   │
│  │  - エラーハンドリング                               │   │
│  │  - ログ記録                                         │   │
│  └──────────────────────────────────────────────────────┘   │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  ビジネスロジック層                                  │   │
│  │  - OpenAI API 連携                                 │   │
│  │  - Excel データ処理                                │   │
│  │  - 分析・操作エンジン                               │   │
│  └──────────────────────────────────────────────────────┘   │
└────────────┬─────────────────────────────────────────────────┘
             │ API呼び出し
             ↓
┌─────────────────────────────────────────────────────────────┐
│              OpenAI API（GPT-4）                             │
│  - テキスト生成                                              │
│  - 自然言語理解                                              │
│  - プロンプトエンジニアリング                                │
└─────────────────────────────────────────────────────────────┘
```

## レイヤー構成

### 1. フロントエンド層（React + TypeScript）

**責務**: ユーザーインターフェースの提供とExcelとの連携

**主要コンポーネント**:

| コンポーネント | 役割 | 主要メソッド |
|-------------|------|----------|
| `TaskPane` | メインコンポーネント | `handleSendMessage()`, `getSelectedData()` |
| `ChatMessage` | メッセージ表示 | `render()` |
| `ChatInput` | ユーザー入力 | `handleSend()`, `handleKeyDown()` |
| `LoadingSpinner` | ローディング表示 | `render()` |

**技術スタック**:
- React 18（UI フレームワーク）
- TypeScript（型安全性）
- Office.js（Excel 連携）
- CSS3（スタイリング）

**通信フロー**:
```
ユーザー入力
  ↓
ChatInput コンポーネント
  ↓
getSelectedData() で Excel データ取得
  ↓
fetch() で API 呼び出し
  ↓
レスポンス処理
  ↓
ChatMessage で表示
  ↓
writeToCell() で Excel に書き込み
```

### 2. バックエンド層（Express + Node.js）

**責務**: API提供、ビジネスロジック実行、外部API連携

**主要モジュール**:

| モジュール | 役割 |
|----------|------|
| `server.js` | Express アプリケーション、ルーティング |
| `openai.js` | OpenAI API 連携、プロンプト処理 |
| `openai-enhanced.js` | 高度な分析機能 |
| `excel-helpers.js` | Excel データ処理ユーティリティ |
| `excel-advanced.js` | 高度な Excel 分析 |
| `middleware.js` | セキュリティ、エラーハンドリング |

**API エンドポイント**:

```
POST /api/chat
├─ 入力: { message, cellData, messageHistory }
├─ 処理: OpenAI API 呼び出し
└─ 出力: { message, action, data }

POST /api/analyze
├─ 入力: { cellData }
├─ 処理: 統計分析
└─ 出力: { statistics, outliers }

GET /health
├─ 入力: なし
├─ 処理: ヘルスチェック
└─ 出力: { status, timestamp }
```

### 3. データ処理層

**Excel ヘルパー関数**:

```javascript
// 統計情報計算
calculateStatistics(values)
  → { sum, average, max, min, median, count }

// 異常値検出
detectOutliers(values)
  → [outlier1, outlier2, ...]

// 空白補完
fillBlankCells(values)
  → [[...], [...], ...]

// ソート
sortData(values, columnIndex, ascending)
  → [[...], [...], ...]

// ピボットテーブル
generatePivotTable(values, rowIndex, colIndex, valueIndex)
  → [[...], [...], ...]

// 回帰分析
performRegression(xValues, yValues)
  → { slope, intercept, rSquared, equation }
```

### 4. AI 処理層

**プロンプトエンジニアリング**:

```
システムプロンプト
  ↓
ユーザーメッセージ + セルデータ
  ↓
OpenAI API (GPT-4)
  ↓
JSON レスポンス
  ↓
パース & アクション実行
  ↓
結果を Excel に反映
```

**アクションタイプ**:

| アクション | 説明 | 実装関数 |
|----------|------|---------|
| `analyze` | 統計分析 | `processAnalyzeAdvanced()` |
| `fill` | 空白補完 | `processFill()` |
| `sort` | ソート | `processSort()` |
| `report` | レポート生成 | `processReportAdvanced()` |
| `chart` | グラフ生成 | `processChartAdvanced()` |
| `pivot` | ピボットテーブル | `processPivot()` |
| `regression` | 回帰分析 | `processRegression()` |
| `correlation` | 相関係数 | `processCorrelation()` |
| `moving_average` | 移動平均 | `processMovingAverage()` |
| `none` | アクションなし | - |

## データフロー

### 1. チャットメッセージ処理フロー

```
[フロントエンド]
  1. ユーザーがメッセージを入力
  2. getSelectedData() で Excel データを取得
  3. fetch() で /api/chat にリクエスト送信
     {
       message: "このデータを分析して",
       cellData: { address: "A1:C10", values: [...] },
       messageHistory: [...]
     }

[バックエンド]
  4. validateInput() でバリデーション
  5. validateAPIKey() で API キー確認
  6. chat() で OpenAI API を呼び出し
  7. formatUserMessage() でプロンプト作成
  8. OpenAI から JSON レスポンスを取得
  9. parseAIResponse() で結果をパース
  10. processAction() で適切な処理を実行
  11. 結果を JSON で返却

[フロントエンド]
  12. レスポンスを受け取る
  13. ChatMessage で AI の返答を表示
  14. action が "write" の場合、writeToCell() で Excel に書き込み
  15. メッセージ履歴に追加
```

### 2. データ分析フロー

```
ユーザー: 「このデータを分析して」
  ↓
Excel データ取得（A1:C10）
  ↓
OpenAI API
  システムプロンプト + セルデータ
  ↓
AI が "analyze" アクションを決定
  ↓
processAnalyzeAdvanced()
  ├─ calculateStatistics() → 統計情報
  ├─ detectOutliers() → 異常値
  └─ calculateZScores() → Z スコア
  ↓
結果を JSON で返却
  ↓
フロントエンド表示 + Excel 書き込み
```

## セキュリティアーキテクチャ

### 認証・認可

```
リクエスト
  ↓
validateAPIKey()
  └─ OPENAI_API_KEY が設定されているか確認
  ↓
CORS チェック
  └─ 許可されたオリジンからのリクエストのみ受け入れ
  ↓
レート制限
  └─ 15分間に100リクエスト以上は拒否
  ↓
入力検証
  └─ メッセージサイズ、セルデータサイズをチェック
```

### データ保護

```
セルデータ
  ↓
メモリ内でのみ処理
  ↓
ログには出力しない
  ↓
API 呼び出し後は破棄
  ↓
永続化なし
```

### ネットワークセキュリティ

```
セキュリティヘッダー
  ├─ X-Content-Type-Options: nosniff
  ├─ X-Frame-Options: DENY
  ├─ X-XSS-Protection: 1; mode=block
  └─ Content-Security-Policy: ...

HTTPS
  └─ 本番環境では必須

CORS
  └─ 信頼できるオリジンのみ許可
```

## パフォーマンス最適化

### フロントエンド

```javascript
// メッセージ履歴の制限
const maxHistory = 10;

// 自動スクロール
useEffect(() => {
  scrollToBottom();
}, [messages]);

// 条件付きレンダリング
{messages.length === 0 ? <EmptyState /> : <Messages />}
```

### バックエンド

```javascript
// リクエストタイムアウト
requestTimeout(30000)

// レート制限
createRateLimiter(15 * 60 * 1000, 100)

// キャッシング（実装予定）
// Redis キャッシュを使用して同じ分析結果を再利用

// バッチ処理
// 複数の操作をまとめて実行
```

### データベース（将来実装）

```
キャッシング層
  ├─ Redis
  └─ メモリ内キャッシュ

データベース
  ├─ PostgreSQL
  ├─ ユーザーデータ
  ├─ 分析履歴
  └─ キャッシュ
```

## エラーハンドリング戦略

### エラー分類

```
クライアント側エラー（4xx）
  ├─ 400: バリデーションエラー
  ├─ 401: 認証エラー
  ├─ 403: 認可エラー
  ├─ 404: リソースなし
  └─ 429: レート制限

サーバー側エラー（5xx）
  ├─ 500: 内部エラー
  └─ 503: サービス利用不可
```

### エラーハンドリングフロー

```
エラー発生
  ↓
errorHandler() ミドルウェア
  ├─ エラータイプを判定
  ├─ 適切なステータスコードを設定
  ├─ エラーメッセージを生成
  └─ ログに記録
  ↓
JSON レスポンスで返却
  ↓
フロントエンド
  ├─ エラーメッセージを表示
  └─ ユーザーに対応を指示
```

## スケーラビリティ

### 水平スケーリング

```
ロードバランサー
  ├─ サーバー1
  ├─ サーバー2
  └─ サーバー3

共有リソース
  ├─ Redis（キャッシュ）
  ├─ PostgreSQL（データベース）
  └─ S3（ファイルストレージ）
```

### 垂直スケーリング

```
メモリ増加
  └─ より多くのリクエストを同時処理

CPU 増加
  └─ 複雑な計算を高速化

ストレージ増加
  └─ より多くのデータを保存
```

## テスト戦略

### ユニットテスト

```javascript
// Excel ヘルパー関数のテスト
test('calculateStatistics', () => {
  const values = [[100], [200], [300]];
  const result = calculateStatistics(values);
  expect(result.sum).toBe('600');
  expect(result.average).toBe('200');
});
```

### 統合テスト

```javascript
// API エンドポイントのテスト
test('POST /api/chat', async () => {
  const response = await request(app)
    .post('/api/chat')
    .send({ message: 'test', cellData: {...} });
  expect(response.status).toBe(200);
});
```

### E2E テスト

```javascript
// ユーザーフロー全体のテスト
test('User can analyze data', async () => {
  // 1. Excel でセルを選択
  // 2. チャットに入力
  // 3. 結果を確認
});
```

## 今後の拡張

### 短期（1-3ヶ月）

- [ ] メッセージ履歴の永続化
- [ ] ユーザー認証の実装
- [ ] より多くの AI モデルへの対応
- [ ] オフライン モード

### 中期（3-6ヶ月）

- [ ] データベースの統合
- [ ] キャッシング層の実装
- [ ] マルチユーザー対応
- [ ] 複雑な分析機能の追加

### 長期（6ヶ月以上）

- [ ] 機械学習モデルの統合
- [ ] リアルタイムコラボレーション
- [ ] カスタム AI モデルのサポート
- [ ] エンタープライズ機能

---

**最終更新**: 2025年10月21日

