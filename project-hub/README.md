# ProjectHub

Project management dashboard with 9-pane grid layout, real-time updates, and plugin system.

## Requirements

- [Node.js](https://nodejs.org/) v18 or later

## Quick Start

### Windows

`start.bat` をダブルクリック（またはコマンドプロンプトで実行）

```cmd
start.bat
```

### macOS / Linux

```bash
node start.mjs
```

ブラウザが自動的に `http://localhost:3457` を開きます。

## Options

```bash
node start.mjs --port 8080     # ポート変更
node start.mjs --no-open       # ブラウザ自動起動を無効化
```

## Configuration

`config.json` でポート、プラグイン設定などをカスタマイズできます。

```json
{
  "port": 3457,
  "openBrowser": true,
  "plugins": {
    "folder-watcher": {
      "watchPaths": ["./watched"],
      "debounceMs": 1000
    }
  }
}
```

## Plugin System

`plugins/` フォルダに `.mjs` ファイルを配置すると自動的にロードされます。

- `folder-watcher.mjs` — フォルダ監視で進捗を自動登録

## Data

プロジェクトデータは `data/projects.json` に自動保存されます。

## Development

```bash
npm install
npm run dev        # 開発サーバー (Vite)
npm run build      # dist/ にビルド
npm run lint       # ESLint
```
