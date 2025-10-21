# デプロイメント ガイド

このドキュメントでは、Excel AI チャットアシスタントを本番環境にデプロイする手順を説明します。

## デプロイ前のチェックリスト

- [ ] すべてのテストが成功している
- [ ] `.env` ファイルが `.gitignore` に含まれている
- [ ] OpenAI APIキーが安全に管理されている
- [ ] HTTPS が設定されている
- [ ] CORS設定が本番ドメインに合わせて更新されている
- [ ] ログレベルが本番環境に合わせて設定されている
- [ ] セキュリティヘッダーが設定されている

## デプロイメント方法

### 方法1：Azure App Service（推奨）

Azure App Service は Microsoft の推奨プラットフォームです。

#### 1.1 前提条件

```bash
# Azure CLI をインストール
# https://docs.microsoft.com/en-us/cli/azure/install-azure-cli

# Azure にログイン
az login
```

#### 1.2 リソースグループの作成

```bash
# リソースグループを作成
az group create \
  --name excel-ai-addon-rg \
  --location japaneast

# App Service プランを作成
az appservice plan create \
  --name excel-ai-addon-plan \
  --resource-group excel-ai-addon-rg \
  --sku B1 \
  --is-linux
```

#### 1.3 Web アプリの作成

```bash
# Web アプリを作成
az webapp create \
  --resource-group excel-ai-addon-rg \
  --plan excel-ai-addon-plan \
  --name excel-ai-addon-app \
  --runtime "node|18"
```

#### 1.4 環境変数の設定

```bash
# 環境変数を設定
az webapp config appsettings set \
  --resource-group excel-ai-addon-rg \
  --name excel-ai-addon-app \
  --settings \
    OPENAI_API_KEY="sk-..." \
    NODE_ENV="production" \
    OPENAI_MODEL="gpt-4" \
    PORT="8080"
```

#### 1.5 デプロイ

```bash
# Git リポジトリを初期化
git init
git add .
git commit -m "Initial commit"

# Azure へデプロイ
az webapp deployment source config-zip \
  --resource-group excel-ai-addon-rg \
  --name excel-ai-addon-app \
  --src excel-ai-addon.zip

# または、Git で直接デプロイ
git remote add azure https://excel-ai-addon-app.scm.azurewebsites.net/excel-ai-addon-app.git
git push azure main
```

#### 1.6 デプロイ確認

```bash
# Web アプリの URL を確認
az webapp show \
  --resource-group excel-ai-addon-rg \
  --name excel-ai-addon-app \
  --query defaultHostName

# ヘルスチェック
curl https://excel-ai-addon-app.azurewebsites.net/health
```

### 方法2：Heroku

#### 2.1 前提条件

```bash
# Heroku CLI をインストール
# https://devcenter.heroku.com/articles/heroku-cli

# Heroku にログイン
heroku login
```

#### 2.2 アプリの作成

```bash
# Heroku アプリを作成
heroku create excel-ai-addon

# または既存のアプリを指定
heroku apps:create excel-ai-addon
```

#### 2.3 環境変数の設定

```bash
# 環境変数を設定
heroku config:set \
  OPENAI_API_KEY="sk-..." \
  NODE_ENV="production" \
  OPENAI_MODEL="gpt-4"
```

#### 2.4 デプロイ

```bash
# Git にリモートを追加
heroku git:remote -a excel-ai-addon

# デプロイ
git push heroku main

# ログを確認
heroku logs --tail
```

### 方法3：AWS Lambda + API Gateway

#### 3.1 前提条件

```bash
# AWS CLI をインストール
# https://aws.amazon.com/jp/cli/

# AWS にログイン
aws configure
```

#### 3.2 Lambda 関数の作成

```bash
# Lambda 関数をパッケージ化
zip -r lambda-function.zip . -x "node_modules/*" ".git/*"

# node_modules をインストール
npm install --production
zip -r lambda-function.zip node_modules/

# Lambda 関数を作成
aws lambda create-function \
  --function-name excel-ai-addon \
  --runtime nodejs18.x \
  --role arn:aws:iam::ACCOUNT_ID:role/lambda-role \
  --handler src/server/server.handler \
  --zip-file fileb://lambda-function.zip
```

#### 3.3 環境変数の設定

```bash
# 環境変数を設定
aws lambda update-function-configuration \
  --function-name excel-ai-addon \
  --environment Variables="{OPENAI_API_KEY=sk-...,NODE_ENV=production}"
```

#### 3.4 API Gateway の設定

```bash
# API Gateway を作成
aws apigateway create-rest-api \
  --name excel-ai-addon-api

# リソースとメソッドを設定
# （詳細は AWS ドキュメントを参照）
```

### 方法4：Google Cloud Run

#### 4.1 前提条件

```bash
# Google Cloud SDK をインストール
# https://cloud.google.com/sdk/docs/install

# Google Cloud にログイン
gcloud auth login

# プロジェクトを設定
gcloud config set project PROJECT_ID
```

#### 4.2 Docker イメージの作成

`Dockerfile` を作成：

```dockerfile
FROM node:18-alpine

WORKDIR /app

COPY package*.json ./
RUN npm install --production

COPY . .

ENV PORT=8080
EXPOSE 8080

CMD ["node", "src/server/server.js"]
```

#### 4.3 イメージをビルドしてデプロイ

```bash
# イメージをビルド
gcloud builds submit --tag gcr.io/PROJECT_ID/excel-ai-addon

# Cloud Run にデプロイ
gcloud run deploy excel-ai-addon \
  --image gcr.io/PROJECT_ID/excel-ai-addon \
  --platform managed \
  --region asia-northeast1 \
  --allow-unauthenticated \
  --set-env-vars OPENAI_API_KEY=sk-...
```

## SSL/TLS 証明書の設定

### Let's Encrypt で証明書を取得

```bash
# certbot をインストール
sudo apt-get install certbot python3-certbot-nginx

# 証明書を取得
sudo certbot certonly --standalone \
  -d your-domain.com \
  -d www.your-domain.com

# 自動更新を設定
sudo systemctl enable certbot.timer
sudo systemctl start certbot.timer
```

### Express で HTTPS を設定

```javascript
const https = require('https');
const fs = require('fs');

const options = {
  key: fs.readFileSync('/etc/letsencrypt/live/your-domain.com/privkey.pem'),
  cert: fs.readFileSync('/etc/letsencrypt/live/your-domain.com/fullchain.pem')
};

https.createServer(options, app).listen(443);
```

## パフォーマンス最適化

### 1. キャッシング

```javascript
// Redis キャッシュを使用
const redis = require('redis');
const client = redis.createClient();

app.get('/api/cache/:key', async (req, res) => {
  const cached = await client.get(req.params.key);
  if (cached) {
    return res.json(JSON.parse(cached));
  }
  // キャッシュミスの場合の処理
});
```

### 2. CDN の設定

```javascript
// CloudFlare CDN を使用
app.use((req, res, next) => {
  res.set('Cache-Control', 'public, max-age=3600');
  next();
});
```

### 3. データベース接続プーリング

```javascript
// コネクションプーリング
const pool = require('pg').Pool;
const pgPool = new pool({
  max: 20,
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 2000,
});
```

## セキュリティ設定

### 1. セキュリティヘッダー

```javascript
const helmet = require('helmet');
app.use(helmet());

// CORS の厳密な設定
app.use(cors({
  origin: 'https://your-domain.com',
  credentials: true,
  methods: ['POST', 'GET'],
  allowedHeaders: ['Content-Type']
}));
```

### 2. レート制限

```javascript
const rateLimit = require('express-rate-limit');

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15分
  max: 100 // 最大100リクエスト
});

app.use('/api/', limiter);
```

### 3. 入力検証

```javascript
const { body, validationResult } = require('express-validator');

app.post('/api/chat',
  body('message').trim().isLength({ min: 1, max: 1000 }),
  (req, res) => {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ errors: errors.array() });
    }
    // 処理を続行
  }
);
```

## モニタリング

### ログ管理

```javascript
const winston = require('winston');

const logger = winston.createLogger({
  level: 'info',
  format: winston.format.json(),
  transports: [
    new winston.transports.File({ filename: 'error.log', level: 'error' }),
    new winston.transports.File({ filename: 'combined.log' })
  ]
});

logger.info('Application started');
```

### エラートラッキング

```javascript
const Sentry = require("@sentry/node");

Sentry.init({ dsn: "YOUR_SENTRY_DSN" });

app.use(Sentry.Handlers.errorHandler());
```

## 本番環境チェックリスト

デプロイ後の確認項目：

- [ ] ヘルスチェックエンドポイントが応答している
- [ ] HTTPS が有効になっている
- [ ] API キーが環境変数で設定されている
- [ ] ログが正常に記録されている
- [ ] エラーハンドリングが機能している
- [ ] CORS が正しく設定されている
- [ ] レート制限が有効になっている
- [ ] バックアップが設定されている
- [ ] 監視とアラートが設定されている

## トラブルシューティング

### デプロイ後にアプリが起動しない

```bash
# ログを確認
heroku logs --tail
# または
az webapp log tail --resource-group excel-ai-addon-rg --name excel-ai-addon-app

# 環境変数を確認
heroku config
# または
az webapp config appsettings list --resource-group excel-ai-addon-rg --name excel-ai-addon-app
```

### API キーエラーが発生する

```bash
# 環境変数が正しく設定されているか確認
echo $OPENAI_API_KEY

# 設定し直す
heroku config:set OPENAI_API_KEY="sk-..."
```

### パフォーマンスが低い

```bash
# メモリ使用量を確認
heroku ps

# スケールアップ
heroku ps:scale web=2
```

## ロールバック

デプロイに問題がある場合：

```bash
# Heroku でロールバック
heroku releases
heroku rollback v123

# Azure でロールバック
az webapp deployment slot swap \
  --resource-group excel-ai-addon-rg \
  --name excel-ai-addon-app \
  --slot staging
```

---

**最終更新**: 2025年10月21日

