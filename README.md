# PPT Finalizer

PowerPoint (`.pptx`) の仕上げをブラウザ内で行うツール。ファイルはサーバーへ送信されません。

- **Web:** https://symmr.github.io/ppt-finalizer/
- **GitHub:** https://github.com/symmr/ppt-finalizer/

画面下部にバージョン（例: `v1.0`）が表示されます。

リポジトリ構成: `docs/index.html`（UI）+ `docs/app.js`（画面の処理）+ `docs/pptx-core.js`（PPTX 解析・変換エンジン、DOM 非依存）+ `docs/version.json`（バージョン）。  
変更履歴は [CHANGELOG.md](./CHANGELOG.md) を参照。

## 機能

- タイトル／本文フォントの統一（既定: Noto Sans JP）
- ファイル分析（サイズ・スライド数・削減見込み）
- フォント分析（使用フォント・埋込ファイルサイズ）
- メディア容量分析（ホバーで画像プレビュー）
- 画像圧縮（表示サイズ基準の ppi。既定 150 ppi、JPEG 再圧縮）
- 孤立メディア削除
- 未使用レイアウト／マスター削除（オプション）
- スピーカーノート削除（オプション）
- ファイルプロパティ削除（オプション）
- 設定の記憶（localStorage）

## 使い方

1. 上記 Web URL を Chrome または Edge で開く
2. `.pptx` をドロップまたは開く
3. タイトル／本文フォントと整理オプションを確認
4. **仕上げてダウンロード** または **仕上げて上書き**

上書き保存は File System Access API 対応ブラウザ（Chrome / Edge）で HTTPS 上のみ利用できます。

## 開発

`docs/pptx-core.js`（PPTX 解析・変換ロジック）に対するユニットテストがあります。GitHub Actions でも `npm test` を実行します。

```sh
npm install
npm test
```

手動確認用に `test/fixtures/smoke-test.pptx`（タイトル/本文フォントが異なる・縮小対象の画像あり・孤立メディアあり・未使用レイアウトあり）を置いています。ブラウザで実際にドロップして動作確認する際に使えます。生成スクリプトは `test/fixtures/build-smoke-pptx.js`（`node test/fixtures/build-smoke-pptx.js` で再生成可能）。

## ライセンス

MIT
