# Changelog

All notable changes to this plugin are documented here. Versioning follows [SemVer](https://semver.org).

## [0.1.0] - 2026-06-28
### Added
- 初回リリース。`pptx-to-md` サブエージェントと処理スクリプト一式を同梱。
- `convert.py`: pptx → Markdown 変換。貼付画像を `<名前>/images/` に抽出し、md の画像参照を実ファイルへの相対パスに書き換え。
- `render_images.py`: 文字・テンプレート装飾・表を除いたスライドを PNG 化。中身が空のスライドはスキップ。
- `extract_images.py`: 埋め込み画像をスライド単位で書き出し（SVG はそのまま、EMF/WMF は PNG 変換）。
- `extract_tables.py`: 表を CSV（UTF-8 BOM 付き）に書き出し。
- `build_index.py`: スライド別 `index.json` と全デッキ横断 `chunks.jsonl`（RAG 取り込み用）を生成。
- `pptx_images.py`: 画像抽出の共有ロジック（markitdown と同じ図形列挙順を再現）。
