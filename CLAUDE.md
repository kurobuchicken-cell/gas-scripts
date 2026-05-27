# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

Google Apps Script (GAS) スクリプトの開発リポジトリ。`clasp` を使ってローカルで開発し、Google のサーバーにデプロイする。

## 開発ツール・セットアップ

### clasp（Google Apps Script CLI）

```bash
# インストール（未インストールの場合）
npm install -g @google/clasp

# Google アカウントへのログイン
clasp login

# 既存プロジェクトのクローン
clasp clone <scriptId>

# 新規プロジェクトの作成
clasp create --title "プロジェクト名" --type standalone
```

### よく使うコマンド

```bash
# ローカルの変更を GAS にプッシュ
clasp push

# GAS の最新をローカルにプル
clasp pull

# ブラウザでスクリプトエディタを開く
clasp open

# デプロイ一覧を確認
clasp deployments

# 新しいバージョンをデプロイ
clasp deploy --description "バージョン説明"

# 実行ログをリアルタイムで表示
clasp logs --watch
```

## プロジェクト構造

```
gas-scripts/
├── .clasp.json          # clasp 設定（scriptId、rootDir 等）
├── appsscript.json      # GAS マニフェスト（タイムゾーン、スコープ等）
├── src/                 # TypeScript ソースファイル（clasp + TypeScript 構成の場合）
│   └── *.ts
└── *.js / *.gs          # GAS スクリプトファイル（JavaScript 直書きの場合）
```

## GAS 開発上の注意点

### 言語・実行環境
- GAS は V8 ランタイムで動作する JavaScript（ES2019 相当）。
- `console.log` の代わりに `Logger.log()` または `console.log()`（V8）を使う。
- 非同期処理（`async/await`、`Promise`）は GAS では使えない。すべて同期処理で書く。

### TypeScript を使う場合
- `clasp` は TypeScript をサポートする（`tsconfig.json` が必要）。
- `@types/google-apps-script` をインストールすると型補完が効く。

```bash
npm install --save-dev @types/google-apps-script typescript
```

### スコープ（OAuth）
- `appsscript.json` の `oauthScopes` に必要な権限を明示する。
- スコープは最小権限の原則で追加する。

### トリガー
- 時間駆動・イベント駆動のトリガーは GAS エディタ上か `ScriptApp.newTrigger()` で設定する。
- `clasp push` だけではトリガーは更新されない。
