# ★yoyakucho-app / 予約帳アプリ

## 概要
訪問歯科の予約・スケジュール管理Webアプリ。
フロントエンドはVite+React、バックエンドはGoogle Apps Script(GAS)。

## 技術スタック
- **Frontend**: React + TypeScript + Vite
- **Backend**: Google Apps Script（`GAS_backend.gs`）
- **Build**: `npm run build` → `dist/` フォルダに出力
- **データソース**: Google Drive（`sync-from-gdrive.ps1` で同期）

## ディレクトリ構成
```
★yoyakucho-app/
├── src/               ← Reactソースコード
├── public/            ← 静的ファイル
├── dist/              ← ビルド成果物（デプロイ用）
├── GAS_backend.gs     ← GASバックエンドロジック
├── sync-from-gdrive.ps1 ← Google Driveから同期するスクリプト
├── vite.config.ts
└── package.json
```

## よく使うコマンド
```bash
# 開発サーバー起動
npm run dev

# 本番ビルド
npm run build

# Google Driveから最新データ同期
powershell -File sync-from-gdrive.ps1
```

## デプロイ手順
1. `npm run build` で dist/ を生成
2. dist/ をデプロイ先にアップロード
   （GitHubPages / VPS / その他 → 要確認）

## 注意事項
- GASのWebアプリURLはGASデプロイ画面で確認
- Google Drive同期はPowerShellスクリプトで実行
