# セットアップ詳細ガイド

このドキュメントでは、Excel AI チャットアシスタントの詳細なセットアップ手順を説明します。

## 前提条件

### システム要件
- **OS**: Windows 10/11, macOS 10.15+, Linux
- **Node.js**: v18.0.0以上
- **npm**: v9.0.0以上（またはpnpm）
- **メモリ**: 最小2GB、推奨4GB以上
- **ディスク容量**: 500MB以上

### ソフトウェア要件
- **Microsoft Excel**: 2016以降（デスクトップ版またはOnline）
- **ブラウザ**: Chrome, Edge, Safari, Firefox（最新版）
- **OpenAI APIキー**: https://platform.openai.com で取得

## ステップ1：OpenAI APIキーの取得

### 1.1 アカウント作成

1. https://platform.openai.com にアクセス
2. 「Sign up」をクリック
3. メールアドレスでアカウント作成（または既存アカウントでログイン）
4. メール確認を完了

### 1.2 APIキーの生成

1. ログイン後、左メニューから「API keys」を選択
2. 「Create new secret key」をクリック
3. キーをコピーして安全に保管
4. **注意**: キーは一度だけ表示されるため、必ずコピーしておく

### 1.3 利用額の設定（重要）

1. 左メニューから「Billing」→「Usage limits」を選択
2. 月額上限を設定（推奨: ¥3,000-5,000）
3. 「Save」をクリック

## ステップ2：プロジェクトのセットアップ

### 2.1 リポジトリのクローン

```bash
# GitHubからクローン
git clone https://github.com/your-username/excel-ai-addon.git
cd excel-ai-addon

# または、ZIPファイルを解凍
unzip excel-ai-addon.zip
cd excel-ai-addon
```

### 2.2 依存関係のインストール

```bash
# npm を使用
npm install

# または pnpm を使用（高速）
pnpm install
```

インストール中に以下のパッケージが導入されます：
- React 18
- Express
- OpenAI SDK
- Office.js
- その他の依存関係

### 2.3 環境変数の設定

```bash
# .env.example をコピー
cp .env.example .env

# .env ファイルを編集
# エディタで .env を開く
```

`.env` ファイルの内容：

```env
# OpenAI API設定（必須）
OPENAI_API_KEY=sk-... # ステップ1で取得したキー

# サーバー設定
PORT=3001
NODE_ENV=development

# CORS設定
CORS_ORIGIN=https://localhost:3000

# API設定
OPENAI_MODEL=gpt-4
MAX_TOKENS=2000
```

**重要な注意事項：**
- `OPENAI_API_KEY` は必ず設定してください
- `.env` ファイルはGitに追加しないでください（`.gitignore` に含まれています）
- APIキーは絶対に他人と共有しないでください

## ステップ3：開発環境での実行

### 3.1 サーバーの起動

```bash
# 方法1：サーバーとクライアントを同時に起動
npm start

# 方法2：別々に起動（推奨）
# ターミナル1でサーバー起動
npm run server

# ターミナル2でクライアント起動
npm run client
```

### 3.2 起動確認

サーバーが正常に起動すると、以下のメッセージが表示されます：

```
========================================
Excel AI アドイン バックエンドサーバー
========================================
ポート: 3001
環境: development
OpenAI API: 設定済み
========================================
```

クライアントが起動すると、ブラウザで `http://localhost:3000` が開きます。

### 3.3 ローカルテスト

```bash
# ブラウザで確認
http://localhost:3000

# サーバーのヘルスチェック
curl http://localhost:3001/health
```

## ステップ4：Excelでのサイドロード

### 4.1 Windows版Excelでのサイドロード

1. **Microsoft Excel を開く**
2. **新しいブックを作成**
3. **「挿入」タブをクリック**
4. **「アドイン」→「マイアドイン」を選択**
5. **「マイアドインの管理」をクリック**
6. **「アップロード マイ アドイン」を選択**
7. **`manifest.xml` ファイルを選択**
   - ファイルパス: `excel-ai-addon/manifest.xml`
8. **「アップロード」をクリック**
9. アドインが表示されたら、クリックして起動

### 4.2 Mac版Excelでのサイドロード

1. **Microsoft Excel を開く**
2. **「挿入」タブをクリック**
3. **「アドイン」→「マイアドイン」を選択**
4. **「マイアドインの管理」をクリック**
5. **「アップロード マイ アドイン」を選択**
6. **`manifest.xml` ファイルを選択**
7. **「アップロード」をクリック**

### 4.3 Excel Online でのサイドロード

1. **Office 365 にログイン**
2. **Excel Online を開く**
3. **「挿入」→「アドイン」→「マイアドイン」を選択**
4. **「アップロード マイ アドイン」を選択**
5. **`manifest.xml` ファイルをアップロード**

### 4.4 manifest.xml の編集

本番環境では、`manifest.xml` を編集してください：

```xml
<!-- 開発環境 -->
<SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>

<!-- 本番環境 -->
<SourceLocation DefaultValue="https://your-domain.com/taskpane.html"/>
```

## ステップ5：動作確認

### 5.1 基本的な動作テスト

1. **Excel でテストデータを作成**
   ```
   A1: 100
   A2: 200
   A3: 300
   ```

2. **セル範囲 A1:A3 を選択**

3. **アドインのチャットに入力**
   ```
   このデータを分析して
   ```

4. **「送信」をクリック**

5. **結果を確認**
   - チャットに分析結果が表示される
   - セルに結果が出力される

### 5.2 トラブルシューティング

#### エラー：「Office.jsの初期化に失敗しました」

**原因**: Office.js が読み込めていない

**解決策**:
```bash
# サーバーを再起動
npm run server

# ブラウザキャッシュをクリア
# Ctrl+Shift+Delete（Windows）または Cmd+Shift+Delete（Mac）
```

#### エラー：「設定が完了していません」

**原因**: OpenAI APIキーが設定されていない

**解決策**:
```bash
# .env ファイルを確認
cat .env

# OPENAI_API_KEY が設定されているか確認
# 設定されていなければ、ステップ1を参照して設定
```

#### エラー：「データが選択されていません」

**原因**: Excelでセルが選択されていない

**解決策**:
- Excelでセル範囲を選択してから、もう一度チャットに入力

#### エラー：「ポート3001が使用中です」

**原因**: 別のプロセスがポート3001を使用している

**解決策**:
```bash
# ポート3001を使用しているプロセスを確認
lsof -i :3001

# 別のポートで起動
PORT=3002 npm run server
```

## ステップ6：開発環境の最適化

### 6.1 ホットリロード設定

開発中の自動リロードを有効化：

```bash
# nodemon で自動リロード
npm run server:dev

# React で自動リロード
npm run client:dev
```

### 6.2 デバッグ設定

ブラウザの開発者ツールを使用：

```bash
# Chrome DevTools を開く
F12 または Ctrl+Shift+I（Windows）
Cmd+Option+I（Mac）

# コンソールでエラーを確認
# Network タブで API呼び出しを確認
```

### 6.3 ログ出力

サーバーのログを詳細に出力：

```javascript
// src/server/server.js に追加
console.log('詳細なログ情報');
```

## ステップ7：本番環境への準備

### 7.1 環境変数の本番設定

本番環境では、以下の環境変数を設定：

```env
NODE_ENV=production
OPENAI_MODEL=gpt-4
CORS_ORIGIN=https://your-domain.com
```

### 7.2 SSL/TLS証明書の設定

本番環境では HTTPS を使用：

```bash
# Let's Encrypt で証明書を取得
certbot certonly --standalone -d your-domain.com
```

### 7.3 デプロイ準備

詳細は [DEPLOY.md](./DEPLOY.md) を参照してください。

## よくある質問（FAQ）

### Q1：複数のExcelファイルで同時に使用できますか？

**A**: はい、複数のExcelファイルで同時に使用できます。各ファイルは独立して動作します。

### Q2：オフラインで使用できますか？

**A**: いいえ、OpenAI APIを使用しているため、インターネット接続が必須です。

### Q3：APIの利用料金はいくらですか？

**A**: GPT-4の場合、1000トークンあたり約¥3です。月額¥3,000-5,000程度に設定することをお勧めします。

### Q4：セルデータは保存されますか？

**A**: いいえ、セルデータはメモリ内でのみ処理され、サーバーに保存されません。

### Q5：複数ユーザーで同時に使用できますか？

**A**: はい、複数ユーザーが同時に使用できます。ただし、OpenAI APIのレート制限に注意してください。

## サポート

問題が発生した場合：

1. [README.md](../README.md) のトラブルシューティングを確認
2. ブラウザの開発者ツールでエラーメッセージを確認
3. サーバーのログを確認
4. OpenAI APIの状態を確認（https://status.openai.com）

---

**最終更新**: 2025年10月21日

