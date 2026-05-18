# ProjectHub Plugins

プラグインを `plugins/` フォルダに配置すると、サーバー起動時に自動ロードされます。

## プラグインの作り方

`.mjs` ファイルまたはフォルダ（`index.mjs` を含む）を作成してください。

```js
// plugins/my-plugin.mjs
export default {
  name: 'my-plugin',
  version: '1.0.0',
  description: '概要',

  async init(ctx) {
    // ctx.store     — データストア (getAll, getById, add, update, delete)
    // ctx.broadcast — SSEイベント送信 (event, data)
    // ctx.config    — config.json の plugins.my-plugin セクション
    // ctx.log       — ロガー
    // ctx.rootDir   — アプリルートパス
  },

  async destroy() {
    // シャットダウン時のクリーンアップ
  },
};
```

## 設定

`config.json` の `plugins` セクションにプラグイン固有の設定を記述します：

```json
{
  "plugins": {
    "my-plugin": {
      "key": "value"
    }
  }
}
```

## 組み込みプラグイン

### folder-watcher

フォルダを監視し、ファイル変更時にプロジェクト進捗を自動更新します。

設定例：
```json
{
  "plugins": {
    "folder-watcher": {
      "watchPaths": ["/Users/you/projects"],
      "debounceMs": 1000
    }
  }
}
```
