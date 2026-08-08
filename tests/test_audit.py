#!/usr/bin/env python3
"""Regression tests for the fail-closed sci-diagram-pptx QA scripts."""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw

REPO_ROOT = Path(__file__).resolve().parents[1]
SKILL_DIR = REPO_ROOT / "skills" / "sci-diagram-pptx"
SCRIPTS_DIR = SKILL_DIR / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import audit_pptx
import compare_render
import overflow_check
import qa_gate
import render_evidence


W, H = 9_144_000, 5_143_500
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path: Path, value: dict) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def relationships(entries: list[tuple[str, str, str]]) -> str:
    body = "".join(f'<Relationship Id="{rid}" Type="{kind}" Target="{target}"/>' for rid, kind, target in entries)
    return f'<Relationships xmlns="{REL}">{body}</Relationships>'


def root_sp_tree(children: str = "") -> str:
    return (
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + children + '</p:spTree>'
    )


def native_node() -> str:
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="node_001"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="914400" y="514350"/><a:ext cx="1828800" cy="1028700"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Node</a:t></a:r></a:p></p:txBody></p:sp>'
    )


def image_fill_tile(name: str, object_id: int, x: int, y: int, width: int, height: int) -> str:
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{object_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )


def make_fixture(base: Path, variant: str = "good") -> tuple[Path, Path, Path, Path]:
    source = base / "source.png"
    image = Image.new("RGB", (1600, 900), "white")
    drawing = ImageDraw.Draw(image)
    drawing.rectangle((160, 90, 480, 270), outline="navy", width=8)
    drawing.text((240, 160), "Node", fill="black")
    image.save(source)
    source_sha = digest(source)
    tile = base / "tile.png"
    Image.new("RGB", (32, 32), (210, 230, 250)).save(tile)

    manifest = base / "source-manifest.json"
    write_json(manifest, {
        "manifest_version": "1.0", "kind": "sci-diagram-pptx-source-manifest", "ok": True,
        "source": {"path": str(source), "sha256": source_sha, "format": "PNG", "width_px": 1600, "height_px": 900, "display_width_px": 1600, "display_height_px": 900},
        "checks": [], "errors": [], "warnings": [],
    })
    plan = base / "scene-plan.json"
    write_json(plan, {
        "schema_version": "1.0", "kind": "sci-diagram-pptx-scene-plan",
        "source": {"path": str(source), "sha256": source_sha, "format": "PNG", "width_px": 1600, "height_px": 900},
        "canvas": {"coordinate_system": "normalized-top-left", "width_inches": 10, "height_inches": 5.625},
        "objects": [{
            "id": "node_001", "type": "node", "source_bbox": {"x": .1, "y": .1, "width": .2, "height": .2},
            "bbox": {"x": .1, "y": .1, "width": .2, "height": .2}, "z_order": 0, "confidence": 1,
            "reconstruction": {"mode": "native-exact", "expected_ooxml_kind": "p:sp"},
            "content": {"text": "Node"}, "style": {"shape_geometry": "rect"},
        }],
        "connections": [], "approvals": {"degradations": []},
    })

    extra_shapes = ""
    slide1_rels = [("rId1", f"{R}/slideLayout", "../slideLayouts/slideLayout1.xml")]
    if variant == "image_fill_tiles":
        half_w, half_h = W // 2, H // 2
        extra_shapes = "".join((
            image_fill_tile("tile_1", 3, 0, 0, half_w, half_h),
            image_fill_tile("tile_2", 4, half_w, 0, half_w, half_h),
            image_fill_tile("tile_3", 5, 0, half_h, half_w, H - half_h),
            image_fill_tile("tile_4", 6, half_w, half_h, half_w, H - half_h),
        ))
        slide1_rels.append(("rId2", f"{R}/image", "../media/tile.png"))

    slide1 = f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{root_sp_tree(native_node() + extra_shapes)}</p:cSld></p:sld>'
    src_rect = '<a:srcRect l="1000"/>' if variant == "crop" else ""
    picture = (
        '<p:pic><p:nvPicPr><p:cNvPr id="2" name="source_reference"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="rId2"/>{src_rect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )
    slide2 = f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{root_sp_tree(picture)}</p:cSld></p:sld>'
    layout = f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{root_sp_tree()}</p:cSld></p:sldLayout>'
    master = f'<p:sldMaster xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{root_sp_tree()}</p:cSld></p:sldMaster>'
    presentation = (
        f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>'
        f'<p:sldSz cx="{W}" cy="{H}"/></p:presentation>'
    )
    pptx = base / "deck.pptx"
    with zipfile.ZipFile(pptx, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        archive.writestr("ppt/presentation.xml", presentation)
        archive.writestr("ppt/_rels/presentation.xml.rels", relationships([
            ("rId1", f"{R}/slide", "slides/slide1.xml"), ("rId2", f"{R}/slide", "slides/slide2.xml"),
        ]))
        archive.writestr("ppt/slides/slide1.xml", slide1)
        archive.writestr("ppt/slides/slide2.xml", slide2)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", relationships(slide1_rels))
        archive.writestr("ppt/slides/_rels/slide2.xml.rels", relationships([
            ("rId1", f"{R}/slideLayout", "../slideLayouts/slideLayout1.xml"), ("rId2", f"{R}/image", "../media/source.png"),
        ]))
        archive.writestr("ppt/slideLayouts/slideLayout1.xml", layout)
        archive.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", relationships([("rId1", f"{R}/slideMaster", "../slideMasters/slideMaster1.xml")]))
        archive.writestr("ppt/slideMasters/slideMaster1.xml", master)
        archive.writestr("ppt/media/source.png", source.read_bytes())
        if variant == "image_fill_tiles":
            archive.writestr("ppt/media/tile.png", tile.read_bytes())
    return pptx, source, manifest, plan


class AuditTests(unittest.TestCase):
    def test_good_package_passes_and_modes_are_counted(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            pptx, _, manifest, plan = make_fixture(Path(directory))
            report = audit_pptx.audit_pptx(pptx, manifest, plan)
            self.assertEqual(report["status"], "PASS", report.get("hard_failures"))
            self.assertEqual(report["plan_expected_counts"]["native_objects"], 1)
            self.assertEqual(report["plan_expected_counts"]["mode_counts"], {"native-exact": 1})

    def test_slide2_crop_is_hard_failure(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            pptx, _, manifest, plan = make_fixture(Path(directory), "crop")
            report = audit_pptx.audit_pptx(pptx, manifest, plan)
            checks = {item["id"]: item for item in report["checks"]}
            self.assertEqual(checks["slide2.reference_no_crop"]["status"], "FAIL")

    def test_shape_geometry_mismatch_is_hard_failure(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pptx, _, manifest, plan = make_fixture(root)
            value = json.loads(plan.read_text(encoding="utf-8"))
            value["objects"][0]["style"]["shape_geometry"] = "ellipse"
            write_json(plan, value)
            report = audit_pptx.audit_pptx(pptx, manifest, plan)
            checks = {item["id"]: item for item in report["checks"]}
            self.assertEqual(checks["slide1.named_object_shape_geometry"]["status"], "FAIL")

    def test_textbox_plan_accepts_rect_lowering_with_text_body(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pptx, _, manifest, plan = make_fixture(root)
            value = json.loads(plan.read_text(encoding="utf-8"))
            value["objects"][0]["style"]["shape_geometry"] = "textbox"
            write_json(plan, value)
            report = audit_pptx.audit_pptx(pptx, manifest, plan)
            checks = {item["id"]: item for item in report["checks"]}
            self.assertEqual(checks["slide1.named_object_shape_geometry"]["status"], "PASS")
            self.assertEqual(report["status"], "PASS", report["hard_failures"])

    def test_linear_formula_shape_does_not_require_omml(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pptx, _, manifest, plan = make_fixture(root)
            value = json.loads(plan.read_text(encoding="utf-8"))
            value["objects"][0]["type"] = "formula-node"
            value["objects"][0]["content"]["math_source"] = "Node"
            write_json(plan, value)
            report = audit_pptx.audit_pptx(pptx, manifest, plan)
            checks = {item["id"]: item for item in report["checks"]}
            self.assertEqual(checks["slide1.omml_equations"]["status"], "PASS")
            self.assertEqual(checks["slide1.math_source_semantics"]["status"], "PASS")
            self.assertEqual(report["status"], "PASS", report["hard_failures"])

    def test_four_quadrant_shape_image_fill_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            pptx, _, manifest, plan = make_fixture(Path(directory), "image_fill_tiles")
            report = audit_pptx.audit_pptx(pptx, manifest, plan)
            checks = {item["id"]: item for item in report["checks"]}
            self.assertEqual(checks["slide1.image_bearers_match_isolated_raster_plan"]["status"], "FAIL")
            self.assertEqual(checks["slide1.no_full_slide_image_or_image_background"]["status"], "FAIL")


class VisualTests(unittest.TestCase):
    def make_image(self, path: Path) -> None:
        image = Image.new("RGB", (800, 450), "white")
        drawing = ImageDraw.Draw(image)
        drawing.rectangle((100, 100, 300, 220), outline="black", width=5)
        drawing.rectangle((500, 250, 700, 370), outline="navy", width=5)
        drawing.line((300, 160, 500, 310), fill="black", width=5)
        image.save(path)

    def test_blank_fails_and_identical_passes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference = root / "reference.png"
            self.make_image(reference)
            blank = root / "blank.png"
            Image.new("RGB", (800, 450), "white").save(blank)
            failed = compare_render.compare_images(reference, blank, root / "blank-report")
            passed = compare_render.compare_images(reference, reference, root / "same-report")
            self.assertEqual(failed["status"], "FAIL")
            self.assertEqual(passed["status"], "PASS")

    def test_large_alignment_fails_but_two_pixels_pass(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference = root / "reference.png"
            self.make_image(reference)
            original = Image.open(reference)
            for shift in (2, 15):
                shifted = Image.new("RGB", original.size, "white")
                shifted.paste(original, (shift, 0))
                shifted_path = root / f"shift-{shift}.png"
                shifted.save(shifted_path)
                report = compare_render.compare_images(reference, shifted_path, root / f"report-{shift}")
                check = next(item for item in report["checks"] if item["id"] == "visual.alignment_magnitude")
                self.assertEqual(check["status"], "PASS" if shift == 2 else "FAIL")


class WrapperTests(unittest.TestCase):
    def test_overflow_wrapper_pass_unknown_and_bad_python(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pptx = root / "fake.pptx"
            pptx.write_bytes(b"pptx")
            passing = root / "passing.py"
            passing.write_text('print("Test passed. No overflow detected.")\n', encoding="utf-8")
            unknown = root / "unknown.py"
            unknown.write_text('print("finished")\n', encoding="utf-8")
            self.assertEqual(overflow_check.build_report(pptx, passing, 10, Path(sys.executable))["status"], "PASS")
            self.assertEqual(overflow_check.build_report(pptx, unknown, 10, Path(sys.executable))["status"], "WARN")
            self.assertEqual(overflow_check.build_report(pptx, passing, 10, root / "missing-python")["status"], "FAIL")

    def test_render_wrapper_uses_fresh_directory_and_helper_python(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pptx, _, _, _ = make_fixture(root)
            helper = root / "fake-render.py"
            helper.write_text(
                'import argparse\nfrom pathlib import Path\nfrom PIL import Image\n'
                'p=argparse.ArgumentParser(); p.add_argument("pptx"); p.add_argument("--output_dir"); a=p.parse_args(); d=Path(a.output_dir); d.mkdir(parents=True,exist_ok=True)\n'
                'Image.new("RGB",(20,10),"white").save(d/"slide-1.png"); Image.new("RGB",(20,10),"black").save(d/"slide-2.png")\n'
                'print("Slides rendered to "+str(d))\n',
                encoding="utf-8",
            )
            report = render_evidence.build_report(pptx, helper, root / "renders", 10, Path(sys.executable))
            self.assertEqual(report["status"], "PASS", report["hard_failures"])
            bad = render_evidence.build_report(pptx, helper, root / "renders-2", 10, root / "missing-python")
            self.assertEqual(bad["status"], "FAIL")


class AggregateTests(unittest.TestCase):
    def make_pass_report(self, kind: str, tool: str, version: str, check_ids: set[str], **extra: object) -> dict:
        return {
            "kind": kind, "tool": tool, "tool_version": version, "status": "PASS",
            "hard_failures": [], "warnings": [],
            "checks": [{"id": check_id, "status": "PASS", "severity": "HARD", "message": "pass"} for check_id in sorted(check_ids)],
            **extra,
        }

    def test_empty_shell_cannot_authorize(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = []
            for label, kind in qa_gate.EXPECTED_REPORT_KINDS.items():
                path = root / f"{label}.json"
                write_json(path, {"kind": kind, "status": "PASS", "hard_failures": [], "warnings": [], "checks": []})
                paths.append((label, path))
            summary = qa_gate.aggregate_reports(paths)
            self.assertFalse(summary["delivery_authorization"])
            self.assertEqual(summary["status"], "FAIL")

    def test_fully_bound_reports_authorize(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.png"
            slide1 = root / "slide-1.png"
            slide2 = root / "slide-2.png"
            Image.new("RGB", (20, 10), "white").save(source)
            Image.new("RGB", (20, 10), "black").save(slide1)
            Image.new("RGB", (20, 10), "blue").save(slide2)
            pptx = root / "final.pptx"
            pptx.write_bytes(b"final-pptx")
            scene_plan = root / "scene-plan.json"
            scene_plan.write_text("{}\n", encoding="utf-8")
            source_sha, pptx_sha, scene_sha = digest(source), digest(pptx), digest(scene_plan)

            scene = self.make_pass_report(
                "sci-diagram-pptx-scene-plan-validation", "sci-diagram-pptx/validate_scene_plan.py", "1.0",
                qa_gate.REPORT_CONTRACTS["scene_plan_report"]["checks"],
                scene_plan_sha256=scene_sha, source_match={"actual_sha256": source_sha, "planned_sha256": source_sha},
            )
            audit = self.make_pass_report(
                "sci-diagram-pptx-pptx-audit", "sci-diagram-pptx/audit_pptx.py", "1.0",
                qa_gate.REPORT_CONTRACTS["pptx_audit"]["checks"],
                input={"source_sha256": source_sha, "scene_plan_sha256": scene_sha, "pptx_sha256": pptx_sha},
            )
            visual = self.make_pass_report(
                "sci-diagram-pptx-visual-report", "sci-diagram-pptx/compare_render.py", "1.0",
                qa_gate.REPORT_CONTRACTS["visual_report"]["checks"],
                input={"reference_sha256": source_sha, "render_sha256": digest(slide1)},
            )
            render = self.make_pass_report(
                "sci-diagram-pptx-render-report", "sci-diagram-pptx/render_evidence.py", "1.0",
                qa_gate.REPORT_CONTRACTS["render_report"]["checks"],
                pptx_sha256=pptx_sha,
                renders=[
                    {"slide_index": 1, "path": str(slide1), "sha256": digest(slide1)},
                    {"slide_index": 2, "path": str(slide2), "sha256": digest(slide2)},
                ],
            )
            overflow = self.make_pass_report(
                "sci-diagram-pptx-overflow-report", "sci-diagram-pptx/overflow_check.py", "1.0",
                qa_gate.REPORT_CONTRACTS["overflow_report"]["checks"], pptx_sha256=pptx_sha, schema_version="1.0",
            )

            paths: dict[str, Path] = {}
            for label, value in (("scene_plan_report", scene), ("pptx_audit", audit), ("visual_report", visual), ("render_report", render), ("overflow_report", overflow)):
                paths[label] = root / f"{label}.json"
                write_json(paths[label], value)
            attestation = {
                "kind": "sci-diagram-pptx-manual-review-attestation", "status": "PASS",
                "source_sha256": source_sha, "pptx_sha256": pptx_sha,
                "full_size_visual_review": True, "overflow_check_passed": True,
                "reviewed_at": "2026-08-08T12:00:00Z", "reviewer": "unit-test",
                "evidence": {
                    "slide_1_render": {"path": str(slide1), "sha256": digest(slide1)},
                    "slide_2_render": {"path": str(slide2), "sha256": digest(slide2)},
                    "overflow_report": {"path": str(paths["overflow_report"]), "sha256": digest(paths["overflow_report"]), "exit_code": 0},
                },
            }
            paths["manual_review_attestation"] = root / "manual.json"
            write_json(paths["manual_review_attestation"], attestation)
            order = ("scene_plan_report", "pptx_audit", "visual_report", "overflow_report", "render_report", "manual_review_attestation")
            summary = qa_gate.aggregate_reports([(label, paths[label]) for label in order])
            self.assertTrue(summary["delivery_authorization"], summary["hard_failures"])
            self.assertEqual(summary["status"], "PASS")


if __name__ == "__main__":
    unittest.main(verbosity=2)
