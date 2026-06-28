#!/usr/bin/env python
"""pptx に埋め込まれた画像 (jpg / png / svg など) を、スライド単位で書き出す。

抽出ロジックは pptx_images モジュールに集約（convert.py の md 参照書き換えと
完全に同じ命名・順序・形式変換を使う）。各スライドを markitdown と同じ順序で走査し、
どのスライドにあったか分かるよう `<名前>_slideNNN_imgM.<ext>` という名前で
`docs_from_agent/<名前>/images/` に保存する。

- SVG はベクタのまま `.svg` で保存。
- EMF/WMF は md で表示できないため soffice で PNG に変換して保存。
- その他（jpg/jpeg/png/gif/bmp/tiff 等）は元の形式のまま保存。
- 同じ素材が複数スライドで使われていれば各スライド分が書き出される。

実行には py313 環境 (python-pptx) と、EMF/WMF 変換のため soffice が必要:
    /opt/conda/envs/py313/bin/python src/pptx_to_md/extract_images.py --all

結果は 1 ファイル 1 行の JSON で stdout に出力する。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pptx import Presentation

import pptx_images


def extract_one(pptx_path: Path, outroot: Path) -> dict:
    result: dict = {"input": str(pptx_path), "outdir": None, "images": [], "count": 0, "ok": False}
    if not pptx_path.is_file():
        result["error"] = "input file not found"
        return result

    try:
        prs = Presentation(str(pptx_path))
    except Exception as exc:
        result["error"] = f"open failed: {exc}"
        return result

    out_dir = outroot / pptx_path.stem / "images"
    _targets, written = pptx_images.extract_slide_images(prs, pptx_path.stem, out_dir)

    result.update(
        outdir=str(out_dir) if written else None,
        images=written,
        count=len(written),
        ok=True,
    )
    return result


def collect_inputs(args: argparse.Namespace) -> list[Path]:
    if args.all:
        return sorted(Path(args.pptx_dir).glob("**/*.pptx"))
    return [Path(p) for p in args.inputs]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="pptx に埋め込まれた画像をスライド単位で書き出す")
    parser.add_argument("inputs", nargs="*", help="対象の .pptx ファイル")
    parser.add_argument("--all", action="store_true", help="--pptx-dir 配下の全 .pptx を対象にする")
    parser.add_argument("--pptx-dir", default="pptx", help="--all のときの探索元 (既定: pptx)")
    parser.add_argument("--outdir", default="docs_from_agent", help="出力ルート (既定: docs_from_agent)")
    args = parser.parse_args(argv)

    inputs = collect_inputs(args)
    if not inputs:
        print("no .pptx files found", file=sys.stderr)
        return 1

    outroot = Path(args.outdir)
    outroot.mkdir(parents=True, exist_ok=True)
    failed = 0
    for pptx_path in inputs:
        res = extract_one(pptx_path, outroot)
        if not res["ok"]:
            failed += 1
        print(json.dumps(res, ensure_ascii=False))

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
