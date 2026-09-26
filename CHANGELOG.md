# Changelog

[PPT Finalizer](https://symmr.github.io/ppt-finalizer/) の変更履歴。  
リポジトリ: [github.com/symmr/ppt-finalizer](https://github.com/symmr/ppt-finalizer/)  
画面下部のバージョン表示は `docs/version.json` を参照。

形式は [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠。

## [1.4.1] - 2026-09-04

### Removed

- Splunk RUM と Session Replay

## [1.4.0] - 2026-09-04

### Added

- Splunk RUM（`@splunk/otel-web` v3.1.0）と Session Replay
- 画像圧縮の対象を図形の画像塗り（`p:sp` + `a:blipFill`）と表セルの画像に拡大
- ファイルサイズの警告（50 MB）と上限（200 MB）
- GitHub Actions で `npm test` を実行

### Changed

- 分析時の画像寸法読み取りをヘッダ（先頭 256 KB）のみに変更
- 仕上げ前に分析用 ZIP キャッシュを解放してから再ロード
- グループに `xfrm` が無くても子画像を走査
- `package.json` の版数を画面表示（`docs/version.json`）と揃える

## [1.3.0] - 2026-09-01

### Added

- 画像圧縮（表示サイズ基準の ppi）。既定 150 ppi、96 / 220 ppi を選択可能
- JPEG 品質（標準 / 高）。PNG / WebP は同一形式のまま縮小
- メディア容量分析に現在画素 → 目標画素を表示
- 仕上げ後サイズの削減見込みに画像圧縮を加算

## [1.2.7] - 2026-08-31

### Changed

- 画面下部のバージョン表記を `v1.2.7` のみに簡略化

## [1.2.6] - 2026-08-31

### Changed

- 仕上げ後サイズの表記を簡略化（削減率のみ表示）

## [1.2.5] - 2026-08-31

### Changed

- ファイル分析に「仕上げ後サイズ」行を追加（削減量・割合を併記）

## [1.2.4] - 2026-08-31

### Changed

- メディア容量分析: プレビュー列を追加（PNG/JPEG/GIF/WebP/BMP/**SVG**）
- サムネイルホバーで拡大プレビュー（従来どおり）

## [1.2.3] - 2026-08-31

### Added

- ファビコン（スライド＋チェックマークのアイコン）

### Changed

- JS のキャッシュバスティング（`version.json` 連動）

## [1.2.2] - 2026-08-31

### Changed

- フォントサンプルをセレクトボックス右側（2 列）に配置
- 処理結果の表示順を整理（概要 → 変更内容）

## [1.2.1] - 2026-08-31

### Fixed

- 右パネルタブ切替が効かない問題（`display: flex` が `[hidden]` を上書きしていた）
- 結果タブの青い ● が仕上げ前から表示される問題
- 仕上げ前に結果タブを開いたとき「仕上げると結果がここに表示されます」と表示

## [1.2.0] - 2026-08-31

### Added

- 右パネルを「分析」「結果」タブに分割（仕上げ後は結果タブを自動表示）
- 削減見込みのプログレスバー表示
- 左カラム下部のアクションボタンを sticky 固定
- タイトル／本文フォントのプレビュー表示（未インストール時は警告）

## [1.1.1] - 2026-08-31

### Fixed

- 右側分析パネルがスクロールできない問題（内部スクロールコンテナを追加）

## [1.1.0] - 2026-08-31

### Changed

- ファイル／フォント／メディア分析と処理結果を右側パネルに表示（2 カラムレイアウト）

## [1.0] - 2026-08-28

初回公開リリース。

### Added

- ブラウザ内 PPTX 仕上げ（ファイルはサーバーへ送信しない）
- タイトル／本文フォントの個別統一（既定: Noto Sans JP）と埋込フォント削除
- ファイル分析（ファイルサイズ・スライド数・削減見込み）
- フォント分析（使用フォント・埋込サイズ・使用箇所）
- メディア容量分析（サイズ順一覧・画像ホバープレビュー）
- 整理オプション: 孤立メディア削除、未使用レイアウト／マスター削除、スピーカーノート削除、ファイルプロパティ削除
- 上書き保存（File System Access API、バックアップ付き）と別名ダウンロード
- フォント選択・整理オプションの設定記憶（localStorage）

[1.4.1]: https://github.com/symmr/ppt-finalizer/compare/v1.4.0...v1.4.1
[1.4.0]: https://github.com/symmr/ppt-finalizer/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/symmr/ppt-finalizer/compare/v1.2.7...v1.3.0
[1.2.7]: https://github.com/symmr/ppt-finalizer/compare/v1.2.6...v1.2.7
[1.2.6]: https://github.com/symmr/ppt-finalizer/compare/v1.2.5...v1.2.6
[1.2.5]: https://github.com/symmr/ppt-finalizer/compare/v1.2.4...v1.2.5
[1.2.4]: https://github.com/symmr/ppt-finalizer/compare/v1.2.3...v1.2.4
[1.2.3]: https://github.com/symmr/ppt-finalizer/compare/v1.2.2...v1.2.3
[1.2.2]: https://github.com/symmr/ppt-finalizer/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/symmr/ppt-finalizer/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/symmr/ppt-finalizer/compare/v1.1.1...v1.2.0
[1.1.1]: https://github.com/symmr/ppt-finalizer/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/symmr/ppt-finalizer/compare/v1.0...v1.1.0
[1.0]: https://github.com/symmr/ppt-finalizer/releases/tag/v1.0
