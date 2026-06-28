# pptx-to-md（Claude Code プラグイン）

PowerPoint (`.pptx`) から情報を抽出し、RAG にそのまま載せられる形で出力する Claude Code プラグインです。`pptx-to-md` サブエージェントと、それが呼び出す Python スクリプト一式を同梱しています。

## できること
`./pptx` 下の `.pptx` を読み込み、`./docs_from_agent/<名前>/` に出力します。

| 機能 | 出力 |
|---|---|
| Markdown 変換（貼付画像を抽出し参照を実パスへ書換） | `<名前>/<名前>.md` |
| 文字・テンプレート・表を除いたスライド画像 | `<名前>/slide_images/` |
| 埋め込み画像（SVG はそのまま、EMF/WMF は PNG 変換） | `<名前>/images/` |
| 表 CSV（UTF-8 BOM 付き） | `<名前>/tables/` |
| スライド別インデックス | `<名前>/index.json` |
| RAG 用チャンク（全デッキ横断・1 行 1 スライド） | `docs_from_agent/chunks.jsonl` |

画像・表のファイル名にはスライド番号 (`slideNNN`) が入り、由来スライドが分かります。

## 前提条件（セットアップ）
このプラグインは外部ツールに依存します。**利用前に各自の環境へインストールしてください。**

1. **Python 3.10 以上**（`python3` が PATH にあること）
2. **Python ライブラリ**:
   ```bash
   python3 -m pip install markitdown python-pptx pymupdf
   ```
3. **LibreOffice (`soffice`)** … スライドのレンダリングと EMF/WMF→PNG 変換に使用。PATH に `soffice` が通っていること。
   - Debian/Ubuntu: `sudo apt-get install libreoffice`
   - macOS: `brew install --cask libreoffice`

> venv を使う場合は、その venv の `python3` が PATH 先頭に来るようにするか、エージェントに使う Python を指示してください。

## インストール（プラグイン本体のみ）
マーケットプレイスは用意していません。ローカルディレクトリとして読み込みます。

```bash
# セッション限定で読み込む
claude --plugin-dir /path/to/pptx-to-md
```

または、skills ディレクトリ配下に置くと `pptx-to-md@skills-dir` として自動ロードされます
（例: `~/.claude/skills/pptx-to-md/` にこのディレクトリを配置）。読み込み後 `/agents` に `pptx-to-md` が現れます。

## 使い方
プロジェクト直下に `pptx/` を作り `.pptx` を置いてから、エージェントに依頼します。

```
pptx-to-md で ./pptx の中身を md と画像と表に変換して、index も作って
```

スクリプトを直接実行することもできます（`${CLAUDE_PLUGIN_ROOT}` は実際のプラグイン設置パス）。

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/convert.py"        --all   # md（＋貼付画像抽出・参照書換）
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/render_images.py"  --all   # 文字除去スライド画像
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/extract_images.py" --all   # 埋め込み画像
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/extract_tables.py" --all   # 表 CSV
python3 "${CLAUDE_PLUGIN_ROOT}/scripts/build_index.py"    --all   # index.json ＋ chunks.jsonl（最後に実行）
```

各スクリプトは 1 ファイル 1 行の JSON を標準出力に出します。`build_index.py` は他の出力を読み取って集約するため、**最後**に実行してください。

## ディレクトリ構成
```
pptx-to-md/
  .claude-plugin/plugin.json   # マニフェスト（name 必須）
  agents/pptx-to-md.md         # サブエージェント定義
  scripts/                     # 処理スクリプト（${CLAUDE_PLUGIN_ROOT}/scripts/ から参照）
    convert.py render_images.py extract_images.py extract_tables.py build_index.py pptx_images.py
  README.md  LICENSE  CHANGELOG.md
```

出力（プロジェクト側）:
```
docs_from_agent/
  chunks.jsonl                 # 全デッキ横断・1行1スライド（RAG 取り込み用）
  <名前>/
    <名前>.md
    index.json
    slide_images/  images/  tables/
```

## chunks.jsonl のフィールド（RAG 取り込み）
1 行 1 スライド。`id`(`<deck>#NNN`) / `deck` / `deck_title` / `source_pptx` / `slide` / `slide_count` /
`title` / `text`（本文。画像参照除去・表 markdown 保持）/ `notes`（発表者ノート）/ `has_content` /
`slide_image` / `images` / `tables`。メディアパスは `docs_from_agent/` からの相対です。

## 既知の制限
- スライド番号などの自動フィールドは画像レンダリングに残ることがあります。
- レンダリングは LibreOffice 経由のため、フォント等の見えが PowerPoint と完全一致しないことがあります。
- 旧 `.ppt`（バイナリ形式）は未対応です（`.pptx` のみ）。

## ライセンス
MIT（`LICENSE` 参照）。
