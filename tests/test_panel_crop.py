#!/usr/bin/env python3

from __future__ import annotations

import json
import sys
import tempfile
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "sci-diagram-pptx" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import panel_crop  # noqa: E402


def main() -> int:
    with tempfile.TemporaryDirectory() as directory:
        work = Path(directory)
        source = work / "source.png"
        output = work / "panel.png"
        record = work / "panel.json"
        image = Image.new("RGB", (12, 10), "white")
        for x in range(4, 9):
            for y in range(2, 7):
                image.putpixel((x, y), (20, 80, 160))
        image.save(source)

        args = [
            str(source),
            "--bbox", "4", "2", "5", "5",
            "--output", str(output),
            "--output-json", str(record),
        ]
        output_log = StringIO()
        with redirect_stdout(output_log), redirect_stderr(output_log):
            assert panel_crop.main(args) == 0
        with Image.open(output) as cropped:
            assert cropped.size == (5, 5)
            assert cropped.getpixel((0, 0)) == (20, 80, 160)
        payload = json.loads(record.read_text(encoding="utf-8"))
        assert payload["bbox"] == [4, 2, 5, 5]
        assert payload["output_size_px"] == [5, 5]
        with redirect_stdout(output_log), redirect_stderr(output_log):
            assert panel_crop.main(args) == 1, "must refuse overwrite"
            assert panel_crop.main([
                str(source), "--bbox", "10", "8", "5", "5", "--output", str(work / "bad.png")
            ]) == 1

    print("panel crop: 4/4 PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
