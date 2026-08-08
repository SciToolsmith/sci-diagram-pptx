#!/usr/bin/env python3

from __future__ import annotations

import base64
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "sci-diagram-pptx" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import check_pptx  # noqa: E402


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
W, H = 9_144_000, 5_143_500
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def rels(entries: list[tuple[str, str, str, bool]]) -> str:
    body = "".join(
        f'<Relationship Id="{rid}" Type="{kind}" Target="{target}"'
        + (' TargetMode="External"' if external else "")
        + "/>"
        for rid, kind, target, external in entries
    )
    return f'<Relationships xmlns="{REL}">{body}</Relationships>'


def tree(children: str = "") -> str:
    return (
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + children + "</p:spTree>"
    )


def font_markup(mode: str) -> str:
    if mode == "missing":
        return ""
    if mode == "partial":
        return '<a:latin typeface="Aptos"/>'
    typeface = "+mn-lt" if mode == "theme" else "Aptos"
    return "".join(
        f'<a:{script} typeface="{typeface}"/>' for script in ("latin", "ea", "cs")
    )


def node(variant: str = "good") -> str:
    font_mode = variant if variant in {"missing", "partial", "theme"} else "explicit"
    fonts = font_markup(font_mode)
    geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    if variant in {"custom_missing_rect", "custom_empty_rect", "custom_with_rect", "custom_no_text"}:
        if variant == "custom_with_rect":
            text_rectangle = '<a:rect l="l" t="t" r="r" b="b"/>'
        elif variant == "custom_empty_rect":
            text_rectangle = "<a:rect/>"
        else:
            text_rectangle = ""
        geometry = (
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
            f'<a:pathLst/>{text_rectangle}</a:custGeom>'
        )

    body_properties = "<a:bodyPr/>"
    if variant == "wrap_none":
        body_properties = '<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>'
    elif variant == "normal_autofit":
        body_properties = "<a:bodyPr><a:normAutofit/></a:bodyPr>"
    elif variant == "shape_autofit":
        body_properties = "<a:bodyPr><a:spAutoFit/></a:bodyPr>"
    elif variant == "invalid_autofit":
        body_properties = "<a:bodyPr><a:normAutofit/><a:spAutoFit/></a:bodyPr>"

    if variant == "custom_no_text":
        runs = ""
    elif variant == "unicode_script":
        runs = f'<a:r><a:rPr>{fonts}</a:rPr><a:t>uₖⁱ</a:t></a:r>'
    elif variant == "native_baseline":
        runs = (
            f'<a:r><a:rPr>{fonts}</a:rPr><a:t>x</a:t></a:r>'
            f'<a:r><a:rPr baseline="30000">{fonts}</a:rPr><a:t>2</a:t></a:r>'
        )
    else:
        runs = f'<a:r><a:rPr>{fonts}</a:rPr><a:t>Node</a:t></a:r>'

    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="node_001"/><p:cNvSpPr/><p:nvPr/>'
        '</p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="514350"/>'
        f'<a:ext cx="1828800" cy="1028700"/></a:xfrm>{geometry}</p:spPr>'
        f'<p:txBody>{body_properties}<a:lstStyle/><a:p>{runs}</a:p></p:txBody></p:sp>'
    )


def office_math(mode: str) -> str:
    if mode == "invalid":
        content = "<m:r><m:t>x+1</m:t></m:r>"
    else:
        inner = "" if mode == "empty" else "<m:r><m:t>x+1</m:t></m:r>"
        content = f"<m:oMath>{inner}</m:oMath>"
    return (
        '<p:extLst><p:ext uri="{office-math-test}">'
        f'<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="{M}">'
        f'<m:oMathPara>{content}</m:oMathPara>'
        '</a14:m></p:ext></p:extLst>'
    )


def picture(object_id: int, rid: str, x: int, y: int, width: int, height: int, crop: bool = False) -> str:
    src_rect = '<a:srcRect l="1000"/>' if crop else ""
    return (
        f'<p:pic><p:nvPicPr><p:cNvPr id="{object_id}" name="picture_{object_id}"/>'
        '<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>'
        f'<a:blip r:embed="{rid}"/>{src_rect}<a:stretch><a:fillRect/></a:stretch>'
        f'</p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect">'
        '<a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )


def make_deck(root: Path, variant: str = "good") -> tuple[Path, Path]:
    source = root / "source.png"
    source.write_bytes(PNG)
    one_slide = variant == "one_slide"

    node_variants = {
        "custom_missing_rect", "custom_empty_rect", "custom_with_rect", "custom_no_text", "font_missing", "font_partial", "font_theme",
        "wrap_none", "normal_autofit", "shape_autofit", "invalid_autofit",
        "unicode_script", "native_baseline",
    }
    node_variant = variant.removeprefix("font_") if variant in {"font_missing", "font_partial", "font_theme"} else variant
    slide1_children = node(node_variant if variant in node_variants else "good")
    if variant in {"office_math_valid", "office_math_empty", "office_math_invalid"}:
        slide1_children += office_math(variant.removeprefix("office_math_"))
    slide1_rels = [("rId1", f"{R}/slideLayout", "../slideLayouts/slideLayout1.xml", False)]
    media: dict[str, bytes] = {}
    if variant == "full_image":
        slide1_children += picture(3, "rId2", 0, 0, W, H)
        slide1_rels.append(("rId2", f"{R}/image", "../media/full.png", False))
        media["ppt/media/full.png"] = PNG
    elif variant == "tiles":
        boxes = [
            (0, 0, W // 2, H // 2), (W // 2, 0, W - W // 2, H // 2),
            (0, H // 2, W // 2, H - H // 2),
            (W // 2, H // 2, W - W // 2, H - H // 2),
        ]
        for index, box in enumerate(boxes, start=2):
            slide1_children += picture(index + 1, f"rId{index}", *box)
            slide1_rels.append((f"rId{index}", f"{R}/image", f"../media/tile{index}.png", False))
            media[f"ppt/media/tile{index}.png"] = PNG
    elif variant == "external_media":
        slide1_rels.append(("rId2", f"{R}/image", "https://example.com/source.png", True))

    slide1 = f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{tree(slide1_children)}</p:cSld></p:sld>'
    slide2 = f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{tree(picture(2, "rId2", 0, 0, W, H, variant == "crop"))}</p:cSld></p:sld>'

    slide_ids = '<p:sldId id="256" r:id="rId1"/>'
    presentation_rels = [("rId1", f"{R}/slide", "slides/slide1.xml", False)]
    if not one_slide:
        slide_ids += '<p:sldId id="257" r:id="rId2"/>'
        presentation_rels.append(("rId2", f"{R}/slide", "slides/slide2.xml", False))
    presentation = (
        f'<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst>{slide_ids}</p:sldIdLst>'
        f'<p:sldSz cx="{W}" cy="{H}"/></p:presentation>'
    )

    embedded_source = b"different" if variant == "mismatch" else PNG
    deck = root / f"{variant}.pptx"
    with zipfile.ZipFile(deck, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        archive.writestr("ppt/presentation.xml", presentation)
        archive.writestr("ppt/_rels/presentation.xml.rels", rels(presentation_rels))
        archive.writestr("ppt/slides/slide1.xml", slide1)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", rels(slide1_rels))
        if not one_slide:
            archive.writestr("ppt/slides/slide2.xml", slide2)
            archive.writestr("ppt/slides/_rels/slide2.xml.rels", rels([
                ("rId1", f"{R}/slideLayout", "../slideLayouts/slideLayout1.xml", False),
                ("rId2", f"{R}/image", "../media/source.png", False),
            ]))
            archive.writestr("ppt/media/source.png", embedded_source)
        for part, value in media.items():
            archive.writestr(part, value)
        if variant == "macro":
            archive.writestr("ppt/vbaProject.bin", b"macro")
    return deck, source


class CheckPptxTests(unittest.TestCase):
    def report(self, variant: str) -> dict[str, object]:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        deck, source = make_deck(Path(directory.name), variant)
        return check_pptx.inspect_package(deck, source)

    def ids(self, report: dict[str, object]) -> set[str]:
        return {str(item["id"]) for item in report["hard_failures"]}  # type: ignore[index]

    def warning_ids(self, report: dict[str, object]) -> set[str]:
        return {str(item["id"]) for item in report["warnings"]}  # type: ignore[index]

    def check(self, report: dict[str, object], identifier: str) -> dict[str, object]:
        for item in report["checks"]:  # type: ignore[union-attr]
            if item["id"] == identifier:
                return item
        self.fail(f"missing check {identifier}")

    def test_good_native_two_slide_package_passes(self) -> None:
        report = self.report("good")
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        self.assertEqual(self.check(report, "slide1.font_inventory")["status"], "PASS")
        self.assertEqual(self.check(report, "slide1.text_layout_inventory")["status"], "PASS")

    def test_text_bearing_custom_geometry_without_text_rectangle_fails(self) -> None:
        for variant in ("custom_missing_rect", "custom_empty_rect"):
            with self.subTest(variant=variant):
                report = self.report(variant)
                self.assertIn("slide1.custom_geometry_text_rect", self.ids(report))
                self.assertIn("slide1.custom_geometry_inventory", self.warning_ids(report))

    def test_custom_geometry_with_text_rectangle_is_inventory_warning_only(self) -> None:
        report = self.report("custom_with_rect")
        self.assertEqual(report["status"], "PASS")
        self.assertNotIn("slide1.custom_geometry_text_rect", self.ids(report))
        self.assertEqual(self.check(report, "slide1.custom_geometry_text_rect")["status"], "PASS")
        self.assertIn("slide1.custom_geometry_inventory", self.warning_ids(report))

    def test_non_text_custom_geometry_without_text_rectangle_is_warning_only(self) -> None:
        report = self.report("custom_no_text")
        self.assertEqual(report["status"], "PASS")
        self.assertNotIn("slide1.custom_geometry_text_rect", self.ids(report))
        self.assertIn("slide1.custom_geometry_inventory", self.warning_ids(report))

    def test_theme_partial_and_missing_fonts_are_inventory_warnings(self) -> None:
        for variant in ("font_theme", "font_partial", "font_missing"):
            with self.subTest(variant=variant):
                report = self.report(variant)
                self.assertEqual(report["status"], "PASS")
                self.assertIn("slide1.font_inventory", self.warning_ids(report))

    def test_wrap_and_dynamic_autofit_are_layout_warnings(self) -> None:
        for variant in ("wrap_none", "normal_autofit", "shape_autofit"):
            with self.subTest(variant=variant):
                report = self.report(variant)
                self.assertEqual(report["status"], "PASS")
                self.assertIn("slide1.text_layout_inventory", self.warning_ids(report))

    def test_multiple_autofit_children_fail(self) -> None:
        report = self.report("invalid_autofit")
        self.assertIn("slide1.text_autofit_exclusive", self.ids(report))

    def test_nonempty_office_math_passes(self) -> None:
        report = self.report("office_math_valid")
        self.assertEqual(report["status"], "PASS")
        item = self.check(report, "slide1.office_math")
        self.assertEqual(item["status"], "PASS")
        self.assertEqual(item["evidence"]["math_object_count"], 1)  # type: ignore[index]

    def test_empty_office_math_fails(self) -> None:
        report = self.report("office_math_empty")
        self.assertIn("slide1.office_math", self.ids(report))

    def test_incomplete_office_math_fails(self) -> None:
        report = self.report("office_math_invalid")
        self.assertIn("slide1.office_math", self.ids(report))

    def test_unicode_scripts_warn_but_native_baseline_does_not(self) -> None:
        unicode_report = self.report("unicode_script")
        self.assertEqual(unicode_report["status"], "PASS")
        self.assertIn("slide1.script_notation", self.warning_ids(unicode_report))

        baseline_report = self.report("native_baseline")
        self.assertEqual(baseline_report["status"], "PASS")
        self.assertNotIn("slide1.script_notation", self.warning_ids(baseline_report))
        item = self.check(baseline_report, "slide1.script_notation")
        self.assertEqual(item["evidence"]["native_baseline_property_count"], 1)  # type: ignore[index]

    def test_full_slide_picture_fails(self) -> None:
        report = self.report("full_image")
        self.assertIn("slide1.not_flattened", self.ids(report))

    def test_picture_tile_mosaic_fails(self) -> None:
        report = self.report("tiles")
        self.assertIn("slide1.not_flattened", self.ids(report))

    def test_reference_crop_fails(self) -> None:
        report = self.report("crop")
        self.assertIn("slide2.source_not_cropped", self.ids(report))

    def test_reference_source_mismatch_fails(self) -> None:
        report = self.report("mismatch")
        self.assertIn("slide2.source_bytes", self.ids(report))

    def test_macro_fails(self) -> None:
        report = self.report("macro")
        self.assertIn("package.no_unsafe_embeds", self.ids(report))

    def test_external_media_fails(self) -> None:
        report = self.report("external_media")
        self.assertIn("package.no_unsafe_embeds", self.ids(report))

    def test_one_slide_native_package_passes_with_warning(self) -> None:
        report = self.report("one_slide")
        self.assertEqual(report["status"], "PASS")
        warning_ids = {str(item["id"]) for item in report["warnings"]}  # type: ignore[index]
        self.assertIn("slide2.source_reference", warning_ids)

    def test_non_zip_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bad.pptx"
            path.write_text("not a pptx", encoding="utf-8")
            report = check_pptx.inspect_package(path)
            self.assertEqual(report["status"], "FAIL")
            self.assertIn("package.readable", self.ids(report))


if __name__ == "__main__":
    unittest.main(verbosity=2)
