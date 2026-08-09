#!/usr/bin/env python3

from __future__ import annotations

import base64
import struct
import sys
import tempfile
import unittest
import warnings
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills" / "sci-diagram-pptx" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import check_pptx  # noqa: E402


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
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


def group(
    children: str,
    width: int = W,
    height: int = H,
    child_width: int | None = None,
    child_height: int | None = None,
    x: int = 0,
    y: int = 0,
    rotation: int = 0,
) -> str:
    child_width = child_width or width
    child_height = child_height or height
    return (
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="group_020"/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>'
        f'<a:xfrm rot="{rotation}"><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{width}" cy="{height}"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="{child_width}" cy="{child_height}"/>'
        f'</a:xfrm></p:grpSpPr>{children}</p:grpSp>'
    )


def content_types(slide_count: int) -> str:
    slide_overrides = "".join(
        '<Override '
        f'PartName="/ppt/slides/slide{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for index in range(1, slide_count + 1)
    )
    return (
        f'<Types xmlns="{CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        f'{slide_overrides}'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '</Types>'
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


def picture(
    object_id: int,
    rid: str,
    x: int,
    y: int,
    width: int,
    height: int,
    crop: bool = False,
    hidden: bool = False,
) -> str:
    src_rect = '<a:srcRect l="1000"/>' if crop else ""
    hidden_attribute = ' hidden="1"' if hidden else ""
    return (
        f'<p:pic><p:nvPicPr><p:cNvPr id="{object_id}" name="picture_{object_id}"{hidden_attribute}/>'
        '<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>'
        f'<a:blip r:embed="{rid}"/>{src_rect}<a:stretch><a:fillRect/></a:stretch>'
        f'</p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect">'
        '<a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )


def connector(
    object_id: int,
    x: int,
    y: int,
    width: int,
    height: int,
) -> str:
    return (
        f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{object_id}" name="connector_{object_id}"/>'
        '<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm>'
        f'<a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/>'
        '</a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '</p:spPr></p:cxnSp>'
    )


def make_deck(root: Path, variant: str = "good") -> tuple[Path, Path]:
    source = root / "source.png"
    source.write_bytes(PNG)
    one_slide = variant.startswith("one_slide")

    node_variants = {
        "custom_missing_rect", "custom_empty_rect", "custom_with_rect", "custom_no_text", "font_missing", "font_partial", "font_theme",
        "wrap_none", "normal_autofit", "shape_autofit", "invalid_autofit",
        "unicode_script", "native_baseline",
    }
    node_variant = variant.removeprefix("font_") if variant in {"font_missing", "font_partial", "font_theme"} else variant
    invalid_native_variants = {
        "one_slide_hidden_native_84_raster",
        "one_slide_zero_native_84_raster",
        "one_slide_offslide_native_84_raster",
        "one_slide_vertical_connector",
    }
    slide1_children = (
        "" if variant in invalid_native_variants
        else node(node_variant if variant in node_variants else "good")
    )
    if variant in {"office_math_valid", "office_math_empty", "office_math_invalid"}:
        slide1_children += office_math(variant.removeprefix("office_math_"))
    slide1_rels = [("rId1", f"{R}/slideLayout", "../slideLayouts/slideLayout1.xml", False)]
    media: dict[str, bytes] = {}
    if variant == "one_slide_embedded_source":
        slide1_children += picture(3, "rId2", W // 10, H // 10, W // 10, H // 10)
        slide1_rels.append(("rId2", f"{R}/image", "../media/embedded-source.png", False))
        media["ppt/media/embedded-source.png"] = PNG
    elif variant in {
        "one_slide_source_inset",
        "one_slide_nonstandard_source_inset",
        "one_slide_hidden_source_inset",
        "one_slide_zero_source_inset",
        "one_slide_full_source_inset",
        "one_slide_four_source_insets",
        "one_slide_source_mosaic",
    }:
        source_target = (
            "../assets/embedded-source.png"
            if variant == "one_slide_nonstandard_source_inset"
            else "../media/embedded-source.png"
        )
        source_part = (
            "ppt/assets/embedded-source.png"
            if variant == "one_slide_nonstandard_source_inset"
            else "ppt/media/embedded-source.png"
        )
        slide1_rels.append(("rId2", f"{R}/image", source_target, False))
        media[source_part] = PNG
        if variant in {"one_slide_four_source_insets", "one_slide_source_mosaic"}:
            if variant == "one_slide_source_mosaic":
                boxes = [
                    (0, 0, W // 2, H // 2),
                    (W // 2, 0, W - W // 2, H // 2),
                    (0, H // 2, W // 2, H - H // 2),
                    (W // 2, H // 2, W - W // 2, H - H // 2),
                ]
            else:
                boxes = [
                    (W * index // 5, H // 10, W // 10, H // 10)
                    for index in range(1, 5)
                ]
            for object_id, box in enumerate(boxes, start=3):
                slide1_children += picture(object_id, "rId2", *box, crop=True)
        else:
            width = 0 if variant == "one_slide_zero_source_inset" else W // 5
            height = H // 5
            x, y = W // 10, H // 10
            if variant == "one_slide_full_source_inset":
                x, y, width, height = 0, 0, W, H
            slide1_children += picture(
                3,
                "rId2",
                x,
                y,
                width,
                height,
                crop=True,
                hidden=variant == "one_slide_hidden_source_inset",
            )
    elif variant in {
        "one_slide_hidden_native_84_raster",
        "one_slide_zero_native_84_raster",
        "one_slide_offslide_native_84_raster",
    }:
        native_shape = node()
        if variant == "one_slide_hidden_native_84_raster":
            native_shape = native_shape.replace(
                '<p:cNvPr id="2" name="node_001"/>',
                '<p:cNvPr id="2" name="node_001" hidden="1"/>',
            )
        elif variant == "one_slide_zero_native_84_raster":
            native_shape = native_shape.replace(
                '<a:ext cx="1828800" cy="1028700"/>',
                '<a:ext cx="0" cy="1028700"/>',
            )
        else:
            native_shape = native_shape.replace(
                '<a:off x="914400" y="514350"/>',
                f'<a:off x="{W + 1000}" y="514350"/>',
            )
        slide1_children += native_shape
        slide1_children += picture(3, "rId2", 0, 0, int(W * 0.84), H)
        slide1_rels.append(("rId2", f"{R}/image", "../media/84-percent.png", False))
        media["ppt/media/84-percent.png"] = PNG
    elif variant == "one_slide_vertical_connector":
        slide1_children += connector(3, W // 2, H // 4, 0, H // 2)
    elif variant == "full_image":
        slide1_children += picture(3, "rId2", 0, 0, W, H)
        slide1_rels.append(("rId2", f"{R}/image", "../media/full.png", False))
        media["ppt/media/full.png"] = PNG
    elif variant == "grouped_full_image":
        child_width, child_height = W // 10, H // 10
        slide1_children += group(
            picture(3, "rId2", 0, 0, child_width, child_height),
            child_width=child_width,
            child_height=child_height,
        )
        slide1_rels.append(("rId2", f"{R}/image", "../media/grouped-full.png", False))
        media["ppt/media/grouped-full.png"] = PNG
    elif variant == "rotated_thin_group_image":
        group_width, group_height = int(W * 0.9), int(H * 0.2)
        slide1_children += group(
            picture(3, "rId2", 0, 0, group_width, group_height),
            width=group_width,
            height=group_height,
            x=int(W * 0.05),
            y=int(H * 0.4),
            rotation=45 * 60000,
        )
        slide1_rels.append(("rId2", f"{R}/image", "../media/rotated-thin.png", False))
        media["ppt/media/rotated-thin.png"] = PNG
    elif variant == "missing_media":
        slide1_children += picture(3, "rId2", W // 10, H // 10, W // 10, H // 10)
        slide1_rels.append(("rId2", f"{R}/image", "../media/missing.png", False))
    elif variant == "wrong_media_type":
        slide1_children += picture(3, "rId2", W // 10, H // 10, W // 10, H // 10)
        slide1_rels.append(("rId2", f"{R}/image", "../media/not-image.png", False))
        media["ppt/media/not-image.png"] = PNG
    elif variant == "missing_blip_relationship":
        slide1_children += picture(3, "rId404", W // 10, H // 10, W // 10, H // 10)
    elif variant == "escaping_relationship":
        slide1_rels.append(("rId2", f"{R}/image", "../../../outside.png", False))
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

    layout_children = ""
    layout_rels: list[tuple[str, str, str, bool]] = []
    master_xml: str | None = None
    master_rels: list[tuple[str, str, str, bool]] = []
    slide_attributes = ' showMasterSp="0"' if variant == "one_slide_master_hidden_image" else ""
    if variant == "one_slide_layout_full_image":
        layout_children = picture(30, "rId2", 0, 0, W, H)
        layout_rels.append(("rId2", f"{R}/image", "../assets/layout.png", False))
        media["ppt/assets/layout.png"] = PNG
    elif variant == "one_slide_layout_source_inset":
        layout_children = picture(30, "rId2", W // 10, H // 10, W // 10, H // 10, crop=True)
        layout_rels.append(("rId2", f"{R}/image", "../assets/layout-source.png", False))
        media["ppt/assets/layout-source.png"] = PNG
    elif variant in {"one_slide_master_full_image", "one_slide_master_hidden_image"}:
        layout_rels.append(("rId1", f"{R}/slideMaster", "../slideMasters/slideMaster1.xml", False))
        master_xml = (
            f'<p:sldMaster xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'
            f'<p:cSld>{tree(picture(40, "rId2", 0, 0, W, H))}</p:cSld></p:sldMaster>'
        )
        master_rels.append(("rId2", f"{R}/image", "../assets/master.png", False))
        media["ppt/assets/master.png"] = PNG

    slide1 = f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"{slide_attributes}><p:cSld>{tree(slide1_children)}</p:cSld></p:sld>'
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
        declared_types = content_types(1 if one_slide else 2)
        if master_xml is not None:
            declared_types = declared_types.replace(
                "</Types>",
                '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
                "</Types>",
            )
        if variant == "macro":
            declared_types = declared_types.replace(
                "</Types>",
                '<Override PartName="/ppt/vbaProject.bin" '
                'ContentType="application/vnd.ms-office.vbaProject"/></Types>',
            )
        elif variant == "embedded_payload":
            declared_types = declared_types.replace(
                "</Types>",
                '<Default Extension="bin" ContentType="application/octet-stream"/></Types>',
            )
        elif variant == "wrong_media_type":
            declared_types = declared_types.replace(
                "</Types>",
                '<Override PartName="/ppt/media/not-image.png" '
                'ContentType="application/xml"/></Types>',
            )
        archive.writestr(
            "[Content_Types].xml",
            f'<Types xmlns="{CT}"/>' if variant == "bad_content_types" else declared_types,
        )
        if variant != "missing_root_relationship":
            archive.writestr("_rels/.rels", rels([
                ("rId1", f"{R}/officeDocument", "ppt/presentation.xml", False),
            ]))
        archive.writestr("ppt/presentation.xml", presentation)
        archive.writestr("ppt/_rels/presentation.xml.rels", rels(presentation_rels))
        archive.writestr("ppt/slides/slide1.xml", slide1)
        archive.writestr("ppt/slides/_rels/slide1.xml.rels", rels(slide1_rels))
        archive.writestr(
            "ppt/slideLayouts/slideLayout1.xml",
            f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{tree(layout_children)}</p:cSld></p:sldLayout>',
        )
        if layout_rels:
            archive.writestr(
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                rels(layout_rels),
            )
        if master_xml is not None:
            archive.writestr("ppt/slideMasters/slideMaster1.xml", master_xml)
            archive.writestr(
                "ppt/slideMasters/_rels/slideMaster1.xml.rels",
                rels(master_rels),
            )
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
        elif variant == "empty_unsafe_directories":
            archive.writestr("ppt/embeddings/", b"")
            archive.writestr("ppt/oleObjects/", b"")
            archive.writestr("ppt/activeX/", b"")
        elif variant == "embedded_payload":
            archive.writestr("ppt/embeddings/", b"")
            archive.writestr("ppt/embeddings/payload.bin", b"embedded payload")
    return deck, source


def rewrite_deck(
    deck: Path,
    replacements: dict[str, str | bytes | None],
    additions: dict[str, str | bytes] | None = None,
) -> None:
    temporary = deck.with_name(deck.name + ".rewrite")
    with zipfile.ZipFile(deck) as source, zipfile.ZipFile(
        temporary, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            replacement = replacements.get(info.filename, source.read(info.filename))
            if replacement is None:
                continue
            target.writestr(info, replacement)
        for name, value in (additions or {}).items():
            target.writestr(name, value)
    temporary.replace(deck)


def mark_first_zip_member_encrypted(deck: Path) -> None:
    payload = bytearray(deck.read_bytes())
    for signature, flag_offset in ((b"PK\x03\x04", 6), (b"PK\x01\x02", 8)):
        position = payload.find(signature)
        if position < 0:
            raise AssertionError(f"missing ZIP signature {signature!r}")
        flags = struct.unpack_from("<H", payload, position + flag_offset)[0]
        struct.pack_into("<H", payload, position + flag_offset, flags | 0x1)
    deck.write_bytes(payload)


class CheckPptxTests(unittest.TestCase):
    def report(
        self,
        variant: str,
        *,
        require_single_slide: bool = False,
        inspect_slide: int = 1,
    ) -> dict[str, object]:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        deck, source = make_deck(Path(directory.name), variant)
        return check_pptx.inspect_package(
            deck,
            source,
            require_single_slide=require_single_slide,
            inspect_slide=inspect_slide,
        )

    def ids(self, report: dict[str, object]) -> set[str]:
        return {str(item["id"]) for item in report["hard_failures"]}  # type: ignore[index]

    def warning_ids(self, report: dict[str, object]) -> set[str]:
        return {str(item["id"]) for item in report["warnings"]}  # type: ignore[index]

    def warning(self, report: dict[str, object], identifier: str) -> dict[str, object]:
        for item in report["warnings"]:  # type: ignore[union-attr]
            if item["id"] == identifier:
                return item
        self.fail(f"missing warning {identifier}")

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

    def test_hidden_zero_or_offslide_shape_does_not_satisfy_native_objects(self) -> None:
        variants = {
            "one_slide_hidden_native_84_raster": "hidden",
            "one_slide_zero_native_84_raster": "zero_dimensions",
            "one_slide_offslide_native_84_raster": "outside_slide",
        }
        for variant, excluded_reason in variants.items():
            with self.subTest(variant=variant):
                report = self.report(variant)
                self.assertIn("slide1.native_objects", self.ids(report))
                self.assertNotIn("slide1.not_flattened", self.ids(report))
                failure = next(
                    item for item in report["hard_failures"]  # type: ignore[union-attr]
                    if item["id"] == "slide1.native_objects"
                )
                self.assertEqual(
                    failure["evidence"]["excluded_objects"][excluded_reason],  # type: ignore[index]
                    1,
                )
                flattened = self.check(report, "slide1.not_flattened")
                self.assertAlmostEqual(
                    flattened["evidence"]["largest_coverage"], 0.84, places=3  # type: ignore[index]
                )

    def test_visible_single_axis_connector_satisfies_native_objects(self) -> None:
        report = self.report("one_slide_vertical_connector")
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        native = self.check(report, "slide1.native_objects")
        self.assertEqual(native["evidence"]["native_shapes"], 0)  # type: ignore[index]
        self.assertEqual(native["evidence"]["connectors"], 1)  # type: ignore[index]

    def test_grouped_full_slide_picture_fails(self) -> None:
        report = self.report("grouped_full_image")
        self.assertIn("slide1.not_flattened", self.ids(report))

    def test_rotated_thin_group_picture_does_not_fail_flattening(self) -> None:
        report = self.report("rotated_thin_group_image")
        self.assertEqual(report["status"], "PASS", report["hard_failures"])

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

    def test_empty_unsafe_directory_entries_do_not_fail(self) -> None:
        report = self.report("empty_unsafe_directories")
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        self.assertEqual(self.check(report, "package.no_unsafe_embeds")["status"], "PASS")

    def test_file_inside_embedding_directory_fails(self) -> None:
        report = self.report("embedded_payload")
        self.assertIn("package.no_unsafe_embeds", self.ids(report))
        failure = next(
            item for item in report["hard_failures"]  # type: ignore[union-attr]
            if item["id"] == "package.no_unsafe_embeds"
        )
        self.assertEqual(
            failure["evidence"]["package_parts"],  # type: ignore[index]
            ["ppt/embeddings/payload.bin"],
        )

    def test_external_media_fails(self) -> None:
        report = self.report("external_media")
        self.assertIn("package.no_unsafe_embeds", self.ids(report))

    def test_internal_ole_relationship_fails_outside_conventional_directories(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            deck, source = make_deck(Path(directory), "one_slide")
            with zipfile.ZipFile(deck) as archive:
                relationship_xml = archive.read(
                    "ppt/slides/_rels/slide1.xml.rels"
                ).decode("utf-8")
                types_xml = archive.read("[Content_Types].xml").decode("utf-8")
            relationship_xml = relationship_xml.replace(
                "</Relationships>",
                f'<Relationship Id="rId9" Type="{R}/oleObject" '
                'Target="../custom/payload.bin"/></Relationships>',
            )
            types_xml = types_xml.replace(
                "</Types>",
                '<Default Extension="bin" ContentType="application/octet-stream"/>'
                "</Types>",
            )
            rewrite_deck(
                deck,
                {
                    "ppt/slides/_rels/slide1.xml.rels": relationship_xml,
                    "[Content_Types].xml": types_xml,
                },
                {"ppt/custom/payload.bin": b"embedded package"},
            )
            report = check_pptx.inspect_package(deck, source)
            self.assertIn("package.no_unsafe_embeds", self.ids(report))
            failure = next(
                item for item in report["hard_failures"]  # type: ignore[union-attr]
                if item["id"] == "package.no_unsafe_embeds"
            )
            self.assertEqual(
                failure["evidence"]["unsafe_internal_relationships"][0]["id"],  # type: ignore[index]
                "rId9",
            )

    def test_unsafe_content_type_and_xml_marker_fail_without_path_hint(self) -> None:
        for mode in ("content_type", "xml_marker"):
            with self.subTest(mode=mode), tempfile.TemporaryDirectory() as directory:
                deck, source = make_deck(Path(directory), "one_slide")
                with zipfile.ZipFile(deck) as archive:
                    types_xml = archive.read("[Content_Types].xml").decode("utf-8")
                    slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
                if mode == "content_type":
                    types_xml = types_xml.replace(
                        "</Types>",
                        '<Override PartName="/ppt/custom/control.bin" '
                        'ContentType="application/vnd.ms-office.activeX"/></Types>',
                    )
                    rewrite_deck(
                        deck,
                        {"[Content_Types].xml": types_xml},
                        {"ppt/custom/control.bin": b"control"},
                    )
                else:
                    slide_xml = slide_xml.replace(
                        "</p:sld>", "<p:controls/></p:sld>"
                    )
                    rewrite_deck(deck, {"ppt/slides/slide1.xml": slide_xml})
                report = check_pptx.inspect_package(deck, source)
                self.assertIn("package.no_unsafe_embeds", self.ids(report))

    def test_missing_internal_media_target_fails_package_readability(self) -> None:
        report = self.report("missing_media")
        self.assertIn("package.readable", self.ids(report))

    def test_missing_blip_relationship_fails_package_readability(self) -> None:
        report = self.report("missing_blip_relationship")
        self.assertIn("package.readable", self.ids(report))

    def test_image_relationship_to_non_image_part_fails_package_readability(self) -> None:
        report = self.report("wrong_media_type")
        self.assertIn("package.readable", self.ids(report))

    def test_escaping_internal_relationship_fails_package_readability(self) -> None:
        report = self.report("escaping_relationship")
        self.assertIn("package.readable", self.ids(report))

    def test_invalid_content_types_fails_package_readability(self) -> None:
        report = self.report("bad_content_types")
        self.assertIn("package.readable", self.ids(report))

    def test_missing_root_office_document_relationship_fails(self) -> None:
        report = self.report("missing_root_relationship")
        self.assertIn("package.readable", self.ids(report))

    def test_duplicate_zip_member_fails_before_ambiguous_lookup(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            deck, _ = make_deck(Path(directory), "one_slide")
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                with zipfile.ZipFile(deck, "a", zipfile.ZIP_DEFLATED) as archive:
                    archive.writestr("ppt/presentation.xml", b"duplicate")
            report = check_pptx.inspect_package(deck)
            self.assertIn("package.unique_members", self.ids(report))

    def test_zip_budget_preflight_has_stable_budget_evidence(self) -> None:
        cases = (
            ({"ZIP_MAX_MEMBERS": 1}, "member_count"),
            ({"ZIP_MAX_ENTRY_UNCOMPRESSED": 10}, "single_member"),
            ({"ZIP_MAX_TOTAL_UNCOMPRESSED": 100}, "total"),
            ({"ZIP_MAX_XML_UNCOMPRESSED": 10}, "xml_member"),
            (
                {"ZIP_RATIO_MIN_UNCOMPRESSED": 1, "ZIP_MAX_COMPRESSION_RATIO": 1.0},
                "compression_ratio",
            ),
        )
        for patched, expected_budget in cases:
            with self.subTest(budget=expected_budget), tempfile.TemporaryDirectory() as directory:
                deck, _ = make_deck(Path(directory), "one_slide")
                with mock.patch.multiple(check_pptx, **patched):
                    report = check_pptx.inspect_package(deck)
                self.assertIn("package.zip_budget", self.ids(report))
                failure = next(
                    item for item in report["hard_failures"]  # type: ignore[union-attr]
                    if item["id"] == "package.zip_budget"
                )
                self.assertEqual(failure["evidence"]["budget"], expected_budget)  # type: ignore[index]

    def test_encrypted_member_and_runtime_error_become_stable_failures(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            deck, _ = make_deck(Path(directory), "one_slide")
            mark_first_zip_member_encrypted(deck)
            report = check_pptx.inspect_package(deck)
            self.assertIn("package.encryption", self.ids(report))

        with tempfile.TemporaryDirectory() as directory:
            deck, _ = make_deck(Path(directory), "one_slide")
            with mock.patch.object(
                zipfile.ZipFile, "testzip", side_effect=RuntimeError("encrypted member")
            ):
                report = check_pptx.inspect_package(deck)
            self.assertIn("package.readable", self.ids(report))

    def test_legacy_one_slide_with_source_warns_missing_reference_slide(self) -> None:
        report = self.report("one_slide")
        self.assertEqual(report["status"], "PASS")
        warning_ids = {str(item["id"]) for item in report["warnings"]}  # type: ignore[index]
        self.assertIn("slide2.source_reference", warning_ids)

    def test_single_slide_contract_accepts_external_source(self) -> None:
        report = self.report("one_slide", require_single_slide=True)
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        self.assertEqual(self.check(report, "delivery.single_slide")["status"], "PASS")
        inventory = self.check(report, "source.raster_inset_inventory")
        self.assertEqual(inventory["status"], "PASS")
        self.assertEqual(inventory["evidence"]["source_usage_count"], 0)  # type: ignore[index]
        companion = self.check(report, "source.external_companion")
        self.assertEqual(companion["status"], "PASS")
        self.assertEqual(companion["evidence"]["sha256"], check_pptx.sha256_bytes(PNG))  # type: ignore[index]
        self.assertNotIn("slide2.source_reference", self.warning_ids(report))

    def test_single_slide_contract_allows_visible_cropped_local_source_inset(self) -> None:
        report = self.report("one_slide_source_inset", require_single_slide=True)
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        inventory = self.check(report, "source.raster_inset_inventory")
        evidence = inventory["evidence"]
        self.assertEqual(evidence["source_usage_count"], 1)  # type: ignore[index]
        self.assertEqual(evidence["qualifying_object_count"], 1)  # type: ignore[index]
        self.assertEqual(evidence["violations"], [])  # type: ignore[index]
        self.assertEqual(
            evidence["matching_media_parts"],  # type: ignore[index]
            ["ppt/media/embedded-source.png"],
        )
        source_object = evidence["objects"][0]  # type: ignore[index]
        self.assertTrue(source_object["src_rect_nonzero"])
        self.assertTrue(source_object["qualifies_individually"])
        self.assertGreater(source_object["visible_coverage"], 0)
        companion = self.check(report, "source.external_companion")
        self.assertEqual(companion["evidence"]["qualifying_source_insets"], 1)  # type: ignore[index]
        self.assertNotIn("source.raster_inset_inventory", self.warning_ids(report))

    def test_source_inventory_resolves_nonstandard_image_part_path(self) -> None:
        report = self.report(
            "one_slide_nonstandard_source_inset", require_single_slide=True
        )
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        inventory = self.check(report, "source.raster_inset_inventory")
        self.assertEqual(
            inventory["evidence"]["matching_image_parts"],  # type: ignore[index]
            ["ppt/assets/embedded-source.png"],
        )

    def test_orphan_image_part_with_exact_source_bytes_fails_single_slide_contract(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            deck, source = make_deck(Path(directory), "one_slide")
            rewrite_deck(
                deck,
                {},
                {"ppt/orphans/unreferenced-source.png": PNG},
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True
            )
            self.assertIn("source.not_embedded", self.ids(report))
            inventory = self.warning(report, "source.raster_inset_inventory")
            self.assertEqual(
                inventory["evidence"]["matching_image_parts"],  # type: ignore[index]
                ["ppt/orphans/unreferenced-source.png"],
            )
            self.assertIn(
                "unreferenced_source_media",
                inventory["evidence"]["violations"],  # type: ignore[index]
            )

    def test_source_image_on_layout_is_not_misclassified_as_local_inset(self) -> None:
        report = self.report(
            "one_slide_layout_source_inset", require_single_slide=True
        )
        self.assertIn("source.not_embedded", self.ids(report))
        inventory = self.warning(report, "source.raster_inset_inventory")
        self.assertIn(
            "source_used_outside_inspected_slide",
            inventory["evidence"]["violations"],  # type: ignore[index]
        )

    def test_layout_and_visible_master_images_participate_in_flattening_check(self) -> None:
        for variant in ("one_slide_layout_full_image", "one_slide_master_full_image"):
            with self.subTest(variant=variant):
                report = self.report(variant)
                self.assertIn("slide1.not_flattened", self.ids(report))
                failure = next(
                    item for item in report["hard_failures"]  # type: ignore[union-attr]
                    if item["id"] == "slide1.not_flattened"
                )
                self.assertGreaterEqual(len(failure["evidence"]["layers"]), 2)  # type: ignore[index]

    def test_show_master_shapes_false_excludes_master_image(self) -> None:
        report = self.report("one_slide_master_hidden_image")
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        flattened = self.check(report, "slide1.not_flattened")
        master = next(
            layer for layer in flattened["evidence"]["layers"]  # type: ignore[index]
            if layer["kind"] == "master"
        )
        self.assertFalse(master["shapes_visible"])
        self.assertEqual(master["raster_objects"], 0)

    def test_single_slide_contract_allows_multiple_small_source_insets(self) -> None:
        report = self.report("one_slide_four_source_insets", require_single_slide=True)
        self.assertEqual(report["status"], "PASS", report["hard_failures"])
        inventory = self.check(report, "source.raster_inset_inventory")
        evidence = inventory["evidence"]
        self.assertEqual(evidence["source_object_count"], 4)  # type: ignore[index]
        self.assertEqual(evidence["qualifying_object_count"], 4)  # type: ignore[index]
        self.assertEqual(evidence["violations"], [])  # type: ignore[index]
        self.assertLess(evidence["union_visible_coverage"], 0.80)  # type: ignore[index]

    def test_single_slide_contract_rejects_embedded_source_bytes(self) -> None:
        report = self.report("one_slide_embedded_source", require_single_slide=True)
        self.assertEqual(report["status"], "FAIL")
        self.assertIn("source.not_embedded", self.ids(report))
        inventory = self.warning(report, "source.raster_inset_inventory")
        self.assertIn("missing_src_rect_crop", inventory["evidence"]["violations"])  # type: ignore[index]
        self.assertNotIn(
            "source.external_companion",
            {str(item["id"]) for item in report["checks"]},  # type: ignore[index]
        )
        self.assertNotIn("slide1.not_flattened", self.ids(report))

    def test_single_slide_contract_rejects_hidden_or_zero_size_source_inset(self) -> None:
        variants = {
            "one_slide_hidden_source_inset": "hidden_source_inset",
            "one_slide_zero_source_inset": "zero_dimensions",
        }
        for variant, expected_violation in variants.items():
            with self.subTest(variant=variant):
                report = self.report(variant, require_single_slide=True)
                self.assertIn("source.not_embedded", self.ids(report))
                inventory = self.warning(report, "source.raster_inset_inventory")
                self.assertIn(expected_violation, inventory["evidence"]["violations"])  # type: ignore[index]

    def test_single_slide_contract_rejects_full_page_source_inset(self) -> None:
        report = self.report("one_slide_full_source_inset", require_single_slide=True)
        self.assertNotIn("source.not_embedded", self.ids(report))
        inventory = self.check(report, "source.raster_inset_inventory")
        self.assertEqual(inventory["evidence"]["largest_visible_coverage"], 1.0)  # type: ignore[index]
        self.assertIn("slide1.not_flattened", self.ids(report))

    def test_single_slide_contract_rejects_source_image_mosaic(self) -> None:
        report = self.report("one_slide_source_mosaic", require_single_slide=True)
        self.assertNotIn("source.not_embedded", self.ids(report))
        inventory = self.check(report, "source.raster_inset_inventory")
        self.assertEqual(inventory["evidence"]["source_object_count"], 4)  # type: ignore[index]
        self.assertEqual(inventory["evidence"]["union_visible_coverage"], 1.0)  # type: ignore[index]
        self.assertIn("slide1.not_flattened", self.ids(report))

    def test_single_slide_contract_rejects_extra_slide_without_source_misclassification(self) -> None:
        report = self.report("mismatch", require_single_slide=True)
        self.assertEqual(report["status"], "FAIL")
        self.assertIn("delivery.single_slide", self.ids(report))
        self.assertNotIn("slide2.source_bytes", self.ids(report))
        self.assertNotIn("slide2.source_reference", self.warning_ids(report))

    def test_parser_accepts_require_single_slide(self) -> None:
        args = check_pptx.build_parser().parse_args([
            "deck.pptx", "--source", "source.png", "--require-single-slide",
            "--build-source", "build.mjs",
        ])
        self.assertTrue(args.require_single_slide)
        self.assertEqual(args.build_source, Path("build.mjs"))

    def test_parser_and_api_select_a_one_based_slide(self) -> None:
        args = check_pptx.build_parser().parse_args(["deck.pptx", "--slide", "2"])
        self.assertEqual(args.slide, 2)

        report = self.report("good", inspect_slide=2)
        self.assertEqual(report["inspect_slide"], 2)
        self.assertIn("slide2.native_objects", self.ids(report))
        self.assertIn("slide2.not_flattened", self.ids(report))
        selection = self.check(report, "inspection.slide_selection")
        self.assertEqual(selection["evidence"]["selected_part"], "ppt/slides/slide2.xml")  # type: ignore[index]

        missing = self.report("one_slide", inspect_slide=2)
        self.assertIn("inspection.slide_selection", self.ids(missing))

    def test_portable_build_source_passes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_text(
                'import { Presentation } from "@oai/artifact-tool";\n'
                'import path from "node:path";\n'
                'console.log(Presentation, path.basename(import.meta.url));\n',
                encoding="utf-8",
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertEqual(report["status"], "PASS", report["hard_failures"])
            self.assertEqual(self.check(report, "build_source.portable")["status"], "PASS")

    def test_pptxgenjs_build_source_passes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_text(
                'import pptxgen from "pptxgenjs";\n'
                'import path from "node:path";\n'
                'console.log(pptxgen, path.basename(import.meta.url));\n',
                encoding="utf-8",
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertEqual(report["status"], "PASS", report["hard_failures"])
            portable = self.check(report, "build_source.portable")
            self.assertEqual(portable["evidence"]["imports"], ["node:path", "pptxgenjs"])  # type: ignore[index]

    def test_build_source_rejects_multiple_authoring_runtimes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_text(
                'import PptxGenJS from "pptxgenjs";\n'
                'import { Presentation } from "@oai/artifact-tool";\n'
                'console.log(PptxGenJS, Presentation);\n',
                encoding="utf-8",
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertEqual(report["status"], "FAIL")
            failure = next(
                item for item in report["hard_failures"]  # type: ignore[union-attr]
                if item["id"] == "build_source.portable"
            )
            self.assertIn("more than one authoring runtime", failure["message"])

    def test_build_source_accepts_multiline_imports_and_safe_node_modules(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_text(
                'import fs from "node:fs";\n'
                'import path from "node:path";\n'
                'import { fileURLToPath } from "node:url";\n'
                'import {\n  Presentation,\n} from "@oai/artifact-tool";\n'
                'console.log(fs, path, fileURLToPath, Presentation);\n',
                encoding="utf-8",
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertEqual(report["status"], "PASS", report["hard_failures"])

    def test_build_source_rejects_fail_closed_escape_patterns(self) -> None:
        cases = (
            'import {\n helper,\n} from "./helper.mjs";\nimport P from "pptxgenjs";\n',
            'import P from "pptxgenjs";\nconst module = import(specifier);\n',
            'import P from "pptxgenjs";\nconst module = import/*gap*/(specifier);\n',
            'import P from "pptxgenjs";\nconst helper = require(name);\n',
            'import P from "pptxgenjs";\nconst helper = require/*gap*/(name);\n',
            'import P from "pptxgenjs";\nconst helper = createRequire(import.meta.url);\n',
            'import P from "pptxgenjs";\nimport cp from "node:child_process";\n',
            'import path from "node:path";\nconsole.log(path);\n',
            'import P from "pptxgenjs";\nconst source = "/root/private/source.png";\n',
            'import P from "pptxgenjs";\nconst source = "/srv/app/source.png";\n',
        )
        for content in cases:
            with self.subTest(content=content), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                deck, source = make_deck(root, "one_slide")
                build = root / "build.mjs"
                build.write_text(content, encoding="utf-8")
                report = check_pptx.inspect_package(
                    deck, source, require_single_slide=True, build_source=build
                )
                self.assertIn("build_source.portable", self.ids(report))

    def test_build_source_rejects_second_static_import_on_same_line(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_text(
                'import P from "pptxgenjs"; import helper from "./helper.mjs";\n',
                encoding="utf-8",
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertIn("build_source.portable", self.ids(report))
            failure = next(
                item for item in report["hard_failures"]  # type: ignore[union-attr]
                if item["id"] == "build_source.portable"
            )
            self.assertIn("./helper.mjs", failure["message"])

    def test_machine_path_and_relative_helper_fail_build_source(self) -> None:
        machine_path = "/" + "Users/example/source.png"
        linux_path = "/" + "home/example/source.png"
        for content in (
            f'const source = "{machine_path}";\n',
            f'const source = "{linux_path}";\n',
            'import helper from "./local-helper.mjs";\n',
            'import fs from "fs";\n',
            'import helper from "pptxgenjs/lib/helper.js";\n',
        ):
            with self.subTest(content=content):
                with tempfile.TemporaryDirectory() as directory:
                    root = Path(directory)
                    deck, source = make_deck(root, "one_slide")
                    build = root / "build.mjs"
                    build.write_text(content, encoding="utf-8")
                    report = check_pptx.inspect_package(
                        deck, source, require_single_slide=True, build_source=build
                    )
                    self.assertIn("build_source.portable", self.ids(report))

    def test_non_utf8_build_source_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_bytes(b"\xff\xfe")
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertIn("build_source.portable", self.ids(report))

    def test_high_confidence_secret_fails_build_source(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            deck, source = make_deck(root, "one_slide")
            build = root / "build.mjs"
            build.write_text(
                'const key = "-----BEGIN PRIVATE KEY-----";\n', encoding="utf-8"
            )
            report = check_pptx.inspect_package(
                deck, source, require_single_slide=True, build_source=build
            )
            self.assertIn("build_source.portable", self.ids(report))

    def test_non_zip_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bad.pptx"
            path.write_text("not a pptx", encoding="utf-8")
            report = check_pptx.inspect_package(path)
            self.assertEqual(report["status"], "FAIL")
            self.assertIn("package.readable", self.ids(report))


if __name__ == "__main__":
    unittest.main(verbosity=2)
