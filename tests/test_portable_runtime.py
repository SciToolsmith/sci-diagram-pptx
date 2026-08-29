#!/usr/bin/env python3

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "sci-diagram-pptx"
PROBE = SKILL / "scripts" / "probe_runtime.mjs"
RENDER = SKILL / "scripts" / "render_pptx.py"
CHECKER = SKILL / "scripts" / "check_pptx.py"
TEMPLATE = SKILL / "assets" / "pptxgenjs-build-template.mjs"
NODE = shutil.which("node")


class PortableRuntimeTest(unittest.TestCase):
    maxDiff = None

    def run_probe(
        self,
        *,
        task_dir: Path | None = None,
        path: str | None = None,
        cwd: Path | None = None,
    ) -> tuple[subprocess.CompletedProcess[str], dict]:
        self.assertIsNotNone(NODE, "node is required")
        env = os.environ.copy()
        if path is not None:
            env["PATH"] = path
        command = [str(NODE), str(PROBE), "--runtime", "pptxgenjs"]
        if task_dir is not None:
            command.extend(["--task-dir", str(task_dir)])
        result = subprocess.run(
            command,
            cwd=cwd or ROOT,
            env=env,
            text=True,
            capture_output=True,
            check=False,
        )
        return result, json.loads(result.stdout)

    def test_probe_reports_ready_portable_runtime(self) -> None:
        result, report = self.run_probe(cwd=ROOT)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(report["requestedRuntime"], "pptxgenjs")
        self.assertEqual(report["selectedRuntime"], "pptxgenjs")
        self.assertEqual(Path(report["taskDir"]), ROOT)
        self.assertTrue(report["ready"])
        self.assertEqual(report["missing"], [])
        self.assertTrue(report["checks"]["node"]["supported"])
        self.assertTrue(report["checks"]["pptxgenjs"]["available"])
        self.assertTrue(report["checks"]["pptxgenjs"]["supported"])
        self.assertEqual(report["checks"]["pptxgenjs"]["detectedVersion"], "4.0.1")
        self.assertEqual(report["checks"]["pptxgenjs"]["scope"], "task-dir")
        self.assertTrue(report["checks"]["libreoffice"]["available"])

    def test_probe_reports_missing_commands_without_writes(self) -> None:
        before = {path.relative_to(ROOT) for path in ROOT.rglob("*") if path.is_file()}
        result, report = self.run_probe(task_dir=ROOT, path="")
        after = {path.relative_to(ROOT) for path in ROOT.rglob("*") if path.is_file()}
        self.assertEqual(result.returncode, 1)
        self.assertFalse(report["ready"])
        self.assertIn("command:libreoffice-or-soffice", report["missing"])
        self.assertIn("command:pdftoppm", report["missing"])
        self.assertIn("command:python3-or-python", report["missing"])
        self.assertEqual(before, after, "runtime probe must be read-only")

    def test_probe_rejects_python_older_than_3_10(self) -> None:
        with tempfile.TemporaryDirectory(prefix="portable-python-") as raw:
            binary_dir = Path(raw) / "bin"
            binary_dir.mkdir()
            for name, version in (
                ("python3", "Python 3.9.18"),
                ("libreoffice", "LibreOffice 24.2"),
                ("pdftoppm", "pdftoppm version 24.02"),
            ):
                executable = binary_dir / name
                executable.write_text(
                    f"#!/bin/sh\nprintf '%s\\n' '{version}'\n",
                    encoding="utf-8",
                )
                executable.chmod(0o755)

            result, report = self.run_probe(
                task_dir=ROOT,
                path=str(binary_dir),
                cwd=ROOT,
            )
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertFalse(report["ready"])
            self.assertTrue(report["checks"]["python"]["available"])
            self.assertFalse(report["checks"]["python"]["supported"])
            self.assertIn("python:>=3.10", report["missing"])

    def test_probe_rejects_unsupported_pptxgenjs_version(self) -> None:
        with tempfile.TemporaryDirectory(prefix="portable-version-") as raw:
            runtime_root = Path(raw) / "runtime"
            task = runtime_root / "tasks" / "task-1"
            package = runtime_root / "node_modules" / "pptxgenjs"
            task.mkdir(parents=True)
            package.mkdir(parents=True)
            (package / "package.json").write_text(
                json.dumps(
                    {
                        "name": "pptxgenjs",
                        "version": "3.12.0",
                        "main": "index.cjs",
                    }
                ),
                encoding="utf-8",
            )
            (package / "index.cjs").write_text("module.exports = {};\n", encoding="utf-8")

            result, report = self.run_probe(task_dir=task, cwd=task)
            self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
            self.assertFalse(report["ready"])
            self.assertTrue(report["checks"]["pptxgenjs"]["available"])
            self.assertFalse(report["checks"]["pptxgenjs"]["supported"])
            self.assertEqual(report["checks"]["pptxgenjs"]["detectedVersion"], "3.12.0")
            self.assertIn("npm:pptxgenjs@4.0.x", report["missing"])

    def test_template_rejects_unconfirmed_exif_orientation(self) -> None:
        with tempfile.TemporaryDirectory(prefix="portable-exif-") as raw:
            runtime_root = Path(raw) / "runtime"
            task = runtime_root / "tasks" / "task-1"
            task.mkdir(parents=True)
            (runtime_root / "node_modules").symlink_to(
                ROOT / "node_modules",
                target_is_directory=True,
            )
            source = Image.new("RGB", (1200, 675), "white")
            source.save(task / "source.png")
            unsafe_build = TEMPLATE.read_text(encoding="utf-8").replace(
                'sourceOrientation: "normalized"',
                'sourceOrientation: "exif-active"',
                1,
            )
            build_source = task / "build.mjs"
            build_source.write_text(unsafe_build, encoding="utf-8")

            result = subprocess.run(
                [str(NODE), str(build_source)],
                cwd=task,
                text=True,
                capture_output=True,
                check=False,
            )
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("requires an orientation-normalized source image", result.stderr)
            self.assertFalse((task / "editable.pptx").exists())

    def test_template_build_crop_arrows_overwrite_and_render(self) -> None:
        with tempfile.TemporaryDirectory(prefix="portable-runtime-") as raw:
            runtime_root = Path(raw) / "runtime"
            task = runtime_root / "tasks" / "task-1"
            task.mkdir(parents=True)
            build_source = task / "build.mjs"
            shutil.copy2(TEMPLATE, build_source)

            source = Image.new("RGB", (1200, 675), "white")
            draw = ImageDraw.Draw(source)
            draw.rectangle((850, 165, 1099, 344), fill="#d8edf8", outline="#225577", width=5)
            draw.line((860, 320, 925, 220, 990, 275, 1080, 185), fill="#d14444", width=8)
            source.save(task / "source.png")

            unavailable, unavailable_report = self.run_probe(task_dir=task, cwd=task)
            self.assertEqual(unavailable.returncode, 1)
            self.assertFalse(unavailable_report["ready"])
            self.assertIn("npm:pptxgenjs", unavailable_report["missing"])

            host_node_modules = ROOT / "node_modules"
            self.assertTrue(host_node_modules.is_dir(), "run npm ci before portable tests")
            (runtime_root / "node_modules").symlink_to(
                host_node_modules,
                target_is_directory=True,
            )
            available, available_report = self.run_probe(task_dir=task, cwd=task)
            self.assertEqual(available.returncode, 0, available.stdout + available.stderr)
            self.assertTrue(available_report["ready"])
            self.assertEqual(
                available_report["checks"]["pptxgenjs"]["detectedVersion"],
                "4.0.1",
            )

            build = subprocess.run(
                [str(NODE), str(build_source)],
                cwd=task,
                text=True,
                capture_output=True,
                check=False,
            )
            self.assertEqual(build.returncode, 0, build.stdout + build.stderr)
            report = json.loads(build.stdout)
            self.assertEqual(report["slides"], 1)
            pptx_path = task / "editable.pptx"
            self.assertTrue(pptx_path.is_file())
            self.assertEqual(
                {path.name for path in task.iterdir()},
                {"source.png", "build.mjs", "editable.pptx"},
                "host dependencies must stay outside the three-file delivery task",
            )

            refused = subprocess.run(
                [str(NODE), str(build_source)],
                cwd=task,
                text=True,
                capture_output=True,
                check=False,
            )
            self.assertNotEqual(refused.returncode, 0)
            self.assertIn("already exists", refused.stderr)

            overwritten = subprocess.run(
                [str(NODE), str(build_source), "--overwrite"],
                cwd=task,
                text=True,
                capture_output=True,
                check=False,
            )
            self.assertEqual(overwritten.returncode, 0, overwritten.stdout + overwritten.stderr)

            with zipfile.ZipFile(pptx_path) as archive:
                names = set(archive.namelist())
                self.assertIn("ppt/slides/slide1.xml", names)
                self.assertNotIn("ppt/slides/slide2.xml", names)
                slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
                self.assertIn("replaceable-image-01", slide_xml)
                self.assertIn("Node:question", slide_xml)
                self.assertIn("Edge:question-to-mechanism", slide_xml)

                root = ET.fromstring(slide_xml)
                ns = {
                    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                }
                crop = root.find(".//a:srcRect", ns)
                self.assertIsNotNone(crop)
                crop_values = {key: int(crop.attrib[key]) for key in ("l", "r", "t", "b")}
                self.assertEqual(
                    crop_values,
                    {"l": 70833, "r": 8333, "t": 24444, "b": 48889},
                    "PptxGenJS crop must encode the requested source-pixel rectangle",
                )

                arrows: dict[str, tuple[str | None, str | None]] = {}
                for shape in root.findall(".//p:sp", ns):
                    name_node = shape.find("./p:nvSpPr/p:cNvPr", ns)
                    if name_node is None:
                        continue
                    name = name_node.attrib.get("name", "")
                    line = shape.find("./p:spPr/a:ln", ns)
                    if line is None:
                        continue
                    head = line.find("./a:headEnd", ns)
                    tail = line.find("./a:tailEnd", ns)
                    arrows[name] = (
                        head.attrib.get("type") if head is not None else None,
                        tail.attrib.get("type") if tail is not None else None,
                    )

                self.assertEqual(
                    arrows["Edge:question-to-mechanism:segment-3"],
                    ("none", "triangle"),
                    "forward from->to must put the arrow at the target/end (tailEnd)",
                )
                self.assertEqual(
                    arrows["Edge:mechanism-outcome-feedback:segment-1"],
                    ("triangle", "triangle"),
                    "bidirectional edges must use both source/headEnd and target/tailEnd",
                )

            check_report_path = task / "check.json"
            checked = subprocess.run(
                [
                    sys.executable,
                    str(CHECKER),
                    str(pptx_path),
                    "--source",
                    str(task / "source.png"),
                    "--build-source",
                    str(build_source),
                    "--require-single-slide",
                    "--pass-index",
                    "1",
                    "--output",
                    str(check_report_path),
                ],
                cwd=task,
                text=True,
                capture_output=True,
                check=False,
            )
            check_report_text = check_report_path.read_text(encoding="utf-8")
            self.assertEqual(
                checked.returncode,
                0,
                checked.stdout + checked.stderr + "\n" + check_report_text,
            )
            check_report = json.loads(check_report_text)
            self.assertEqual(check_report["status"], "PASS")
            self.assertEqual(check_report["decision"]["next_action"], "deliver")
            inset_check = next(
                item for item in check_report["checks"]
                if item["id"] == "source.raster_inset_inventory"
            )
            self.assertEqual(inset_check["evidence"]["source_usage_count"], 1)
            self.assertEqual(inset_check["evidence"]["qualifying_object_count"], 1)

            render_dir = task / "render"
            rendered = subprocess.run(
                [
                    sys.executable,
                    str(RENDER),
                    str(pptx_path),
                    "--output-dir",
                    str(render_dir),
                    "--dpi",
                    "96",
                ],
                cwd=task,
                text=True,
                capture_output=True,
                check=False,
            )
            self.assertEqual(rendered.returncode, 0, rendered.stdout + rendered.stderr)
            render_report = json.loads(rendered.stdout)
            self.assertTrue(render_report["ready"])
            self.assertTrue(render_report["isolatedLibreOfficeProfile"])
            self.assertEqual(len(render_report["slides"]), 1)
            slide_png = Path(render_report["slides"][0])
            self.assertTrue(slide_png.is_file())
            with Image.open(slide_png) as image:
                self.assertGreater(image.width, 500)
                self.assertGreater(image.height, 300)

            self.assertNotIn(str(ROOT), build_source.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
