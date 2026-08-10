/**
 * Editable reconstruction of 乳酸菌小分子肽改善胰岛细胞损伤机制图.png
 * Runtime: @oai/artifact-tool 2.8.39 (Node.js 24.14.0)
 * Run: node build.mjs
 * The script expects source.png and writes editable.pptx beside itself.
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE = path.join(HERE, "source.png");
const OUTPUT = path.join(HERE, "editable.pptx");

const SOURCE_WIDTH = 1472;
const SOURCE_HEIGHT = 1068;
const SLIDE_WIDTH = 1472;
const SLIDE_HEIGHT = 1068;

const SONG = "Songti SC";
const LATIN = "Times New Roman";

const COLORS = {
  background: "#FCFBF8",
  title: "#0E88C0",
  titleBorder: "#006D99",
  lightBlue: "#BBD9F2",
  lightBlueBorder: "#6C9FC5",
  paleBlue: "#F2FAFC",
  relationBorder: "#6797B8",
  cream: "#FFF9E9",
  creamBorder: "#D0C5A6",
  mint: "#CEE7DA",
  mintBorder: "#79AA97",
  yellow: "#FFF687",
  yellowBorder: "#D1B929",
  edge: "#19567E",
  dash: "#4A4A4A",
  dashBlue: "#194C86",
  stageArrow: "#ABD7F6",
  stageArrowBorder: "#4D86AF",
};

// Source-derived semantic map. Bounds are [x, y, width, height] in source pixels.
const NODES = [
  { id: "section_left", text: "", type: "semanticRegion", bounds: [0, 0, 410, 664], draw: false },
  { id: "section_center", text: "", type: "semanticRegion", bounds: [411, 0, 491, 664], draw: false },
  { id: "section_right", text: "", type: "semanticRegion", bounds: [902, 0, 365, 664], draw: false },
  { id: "section_summary", text: "", type: "semanticRegion", bounds: [1268, 0, 204, 664], draw: false },
  { id: "section_bottom", text: "", type: "semanticRegion", bounds: [0, 665, 1472, 403], draw: false },

  // The nonstandard phrase “谷胱胺间腺质” is transcribed literally from the source pixels.
  { id: "left_title", text: "L. plantarum SCS2胞内小分子肽调\n控铁死亡及谷胱胺间腺质氧化损伤\n的作用研究", type: "title", bounds: [18, 13, 379, 98], style: "title", fontSize: 22, rich: true },
  { id: "left_peptide", text: "L. plantarum SCS2小分子肽", type: "process", bounds: [72, 132, 270, 44], style: "blue", fontSize: 20, rich: true },
  { id: "left_damage_group", text: "", type: "group", bounds: [35, 196, 344, 61], style: "dashedGray" },
  { id: "left_damage_model", text: "建立脂质稳态损伤\n物模型", type: "process", bounds: [52, 203, 167, 47], style: "blue", fontSize: 18 },
  { id: "left_min6", text: "min6细胞", type: "process", bounds: [245, 204, 115, 46], style: "blue", fontSize: 19 },
  { id: "left_assay_group", text: "", type: "group", bounds: [34, 316, 174, 187], style: "dashedBlue" },
  { id: "left_rtqpcr", text: "RT-qPCR", type: "assay", bounds: [54, 332, 137, 31], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "left_wb", text: "WB", type: "assay", bounds: [54, 372, 137, 31], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "left_elisa", text: "ELISA", type: "assay", bounds: [54, 412, 137, 30], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "left_iron", text: "铁离子比色法", type: "assay", bounds: [54, 451, 137, 31], style: "mint", fontSize: 18 },
  { id: "left_lcms", text: "LC-MS/MS", type: "assay", bounds: [262, 390, 116, 32], style: "mint", fontSize: 18, typeface: LATIN },
  { id: "left_effect", text: "明确L.plantarum\nSCS2胞内小分子肽\n是否具有改善胰岛β\n细胞氧化损伤作用", type: "conclusion", bounds: [12, 531, 220, 118], style: "yellow", fontSize: 20, rich: true, bold: true },
  { id: "left_info", text: "获取理论\n分子量、\n氨基酸序\n列等信息", type: "conclusion", bounds: [256, 535, 127, 114], style: "yellow", fontSize: 20, bold: true },

  { id: "center_title", text: "L. plantarum SCS2胞内小分子肽通\n过lncRNA MEG3调控Nrf2抑制胰\n岛β细胞铁死亡的机制研究", type: "title", bounds: [466, 13, 378, 97], style: "title", fontSize: 21, rich: true },
  { id: "center_rna", text: "RNA pull down实验\nRIP实验\n双荧光素酶报告实验", type: "experiment", bounds: [445, 129, 198, 73], style: "blue", fontSize: 18, bold: true },
  { id: "center_plasmid", text: "lncRNA MEG3 过\n表达/表达缺损质粒\n建和转染", type: "experiment", bounds: [670, 128, 198, 74], style: "blue", fontSize: 18, bold: true },
  { id: "center_relation", text: "明确lncRNA MEG3、 miR-93和Nrf2之间的靶向调控关系", type: "finding", bounds: [422, 213, 466, 54], style: "relation", fontSize: 18, bold: true },
  { id: "center_peptide", text: "L. plantarum SCS2小分子肽", type: "process", bounds: [490, 286, 329, 45], style: "blue", fontSize: 20, rich: true },
  { id: "center_min6", text: "min6细胞", type: "process", bounds: [598, 346, 114, 37], style: "blue", fontSize: 19 },
  { id: "center_meg", text: "沉默/过表达\nlncRNA MEG3", type: "perturbation", bounds: [451, 413, 134, 49], style: "cream", fontSize: 17, bold: true },
  { id: "center_ncoa4", text: "沉默/过表达\nNCOA4", type: "perturbation", bounds: [734, 413, 126, 49], style: "cream", fontSize: 17, bold: true },
  { id: "center_assay_group", text: "", type: "group", bounds: [435, 480, 435, 67], style: "dashedGray" },
  { id: "center_elisa", text: "ELISA\nRT-qPCR", type: "assay", bounds: [452, 490, 102, 46], style: "mint", fontSize: 17, typeface: LATIN },
  { id: "center_wb", text: "WB\n免疫荧光", type: "assay", bounds: [611, 490, 99, 46], style: "mint", fontSize: 17 },
  { id: "center_iron", text: "铁离子比\n色法", type: "assay", bounds: [764, 490, 98, 47], style: "mint", fontSize: 17 },
  { id: "center_conclusion", text: "明确L.plantarum SCS2胞内小分子肽是否通过\nlncRNA MEG3调控Nrf2抑制胰岛β细胞铁死亡", type: "conclusion", bounds: [427, 564, 457, 81], style: "yellow", fontSize: 19, rich: true, bold: true },

  { id: "right_title", text: "L. plantarum SCS2胞内小分子\n肽对lncRNA MEG3启动子去\n甲基化的作用研究", type: "title", bounds: [918, 13, 330, 98], style: "title", fontSize: 21, rich: true },
  { id: "right_peptide1", text: "L. plantarum\nSCS2小分子肽", type: "process", bounds: [997, 138, 171, 52], style: "blue", fontSize: 19, rich: true },
  { id: "right_damage_group", text: "", type: "group", bounds: [916, 216, 332, 72], style: "dashedGray" },
  { id: "right_min6", text: "min6细胞", type: "process", bounds: [929, 226, 110, 50], style: "blue", fontSize: 19 },
  { id: "right_damage_model", text: "建立胰岛β细胞损\n的模型", type: "process", bounds: [1070, 226, 165, 50], style: "blue", fontSize: 18 },
  { id: "right_peptide2", text: "L. plantarum SCS2胞\n内小分子肽", type: "process", bounds: [974, 316, 216, 55], style: "blue", fontSize: 18, rich: true },
  { id: "right_assay_group", text: "", type: "group", bounds: [947, 397, 268, 114], style: "dashedGray" },
  { id: "right_msp", text: "MSP", type: "assay", bounds: [971, 408, 100, 30], style: "mint", fontSize: 18, typeface: LATIN },
  { id: "right_chip", text: "ChIP", type: "assay", bounds: [1097, 408, 97, 30], style: "mint", fontSize: 18, typeface: LATIN },
  { id: "right_elisa", text: "ELISA\nRT-qPCR", type: "assay", bounds: [971, 447, 100, 51], style: "mint", fontSize: 17, typeface: LATIN },
  { id: "right_wb", text: "WB\n免疫荧光", type: "assay", bounds: [1097, 447, 97, 51], style: "mint", fontSize: 17 },
  { id: "right_conclusion", text: "明确L.plantarum SCS2胞内小分\n子肽是否通过参与lncRNA MEG3\n启动子去甲基化调控其表达", type: "conclusion", bounds: [920, 547, 328, 101], style: "yellow", fontSize: 19, rich: true, bold: true },
  { id: "top_summary", text: "阐明\nL.plantarum\nSCS2改善\n胰岛β细胞\n氧化损伤的\n关键物质和\n作用机制", type: "summary", bounds: [1294, 184, 145, 419], style: "pale", fontSize: 21, rich: true, bold: true },

  { id: "bottom_title", text: "L. plantarum SCS2胞内小分\n子肽改善胰岛β细胞氧化损\n伤对胰防控糖尿病的影响\n研究", type: "title", bounds: [39, 778, 297, 143], style: "title", fontSize: 22, rich: true },
  { id: "bottom_model", text: "", type: "modelPanel", bounds: [401, 786, 199, 144], style: "blueLarge" },
  { id: "bottom_model_label", text: "建立2型糖尿病动\n物模型", type: "label", bounds: [416, 836, 171, 69], style: "textOnly", fontSize: 21 },
  { id: "bottom_assay_group", text: "", type: "group", bounds: [670, 678, 304, 363], style: "dashedBlueLarge" },
  { id: "bottom_msp", text: "MSP", type: "assay", bounds: [702, 732, 112, 32], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "bottom_chip", text: "ChIP", type: "assay", bounds: [835, 731, 109, 33], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "bottom_rtqpcr", text: "RT-qPCR", type: "assay", bounds: [701, 821, 113, 32], style: "mint", fontSize: 18, typeface: LATIN },
  { id: "bottom_wb", text: "WB", type: "assay", bounds: [834, 821, 110, 32], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "bottom_elisa", text: "ELISA", type: "assay", bounds: [702, 912, 112, 32], style: "mint", fontSize: 19, typeface: LATIN },
  { id: "bottom_if", text: "免疫荧光", type: "assay", bounds: [834, 912, 111, 32], style: "mint", fontSize: 18 },
  { id: "bottom_iron", text: "铁离子比色法", type: "assay", bounds: [701, 985, 244, 32], style: "mint", fontSize: 18 },
  { id: "bottom_result", text: "明确L.plantarum SCS2胞内小分子肽改\n善胰岛β细胞氧化损伤对2型糖尿病的\n预防作用", type: "conclusion", bounds: [1040, 775, 389, 153], style: "paleLeft", fontSize: 19, rich: true, bold: true },
];

// Direction is always source -> target. Route records the source-observed connector intent.
const EDGES = [
  { id: "stage_left_center", from: "section_left", to: "section_center", direction: "forward", route: "right block arrow", lineStyle: "solid", render: "stageArrow", frame: [394, 359, 35, 53] },
  { id: "stage_center_right", from: "section_center", to: "section_right", direction: "forward", route: "right block arrow", lineStyle: "solid", render: "stageArrow", frame: [884, 359, 35, 53] },
  { id: "stage_right_summary", from: "section_right", to: "top_summary", direction: "forward", route: "right block arrow", lineStyle: "solid", render: "stageArrow", frame: [1237, 359, 53, 53] },
  { id: "left_title_to_peptide", from: "left_title", to: "left_peptide", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "left_peptide_to_model_group", from: "left_peptide", to: "left_damage_group", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "left_model_to_min6", from: "left_damage_model", to: "left_min6", direction: "forward", route: { kind: "straight", fromSide: "right", toSide: "left" }, lineStyle: "solid" },
  { id: "left_model_group_to_assays", from: "left_damage_group", to: "left_assay_group", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "left_model_group_to_lcms", from: "left_damage_group", to: "left_lcms", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "left_assays_to_effect", from: "left_assay_group", to: "left_effect", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "left_lcms_to_info", from: "left_lcms", to: "left_info", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },

  { id: "center_title_to_rna", from: "center_title", to: "center_rna", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_title_to_plasmid", from: "center_title", to: "center_plasmid", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_rna_to_relation", from: "center_rna", to: "center_relation", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_plasmid_to_relation", from: "center_plasmid", to: "center_relation", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_relation_to_peptide", from: "center_relation", to: "center_peptide", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_peptide_to_min6", from: "center_peptide", to: "center_min6", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_min6_to_meg", from: "center_min6", to: "center_meg", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_min6_to_ncoa4", from: "center_min6", to: "center_ncoa4", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_meg_to_assays", from: "center_meg", to: "center_assay_group", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_ncoa4_to_assays", from: "center_ncoa4", to: "center_assay_group", direction: "forward", route: { kind: "elbow", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "center_assays_to_conclusion", from: "center_assay_group", to: "center_conclusion", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },

  { id: "right_title_to_peptide", from: "right_title", to: "right_peptide1", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "right_peptide_to_model_group", from: "right_peptide1", to: "right_damage_group", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "right_min6_to_model", from: "right_min6", to: "right_damage_model", direction: "forward", route: { kind: "straight", fromSide: "right", toSide: "left" }, lineStyle: "solid" },
  { id: "right_model_group_to_peptide", from: "right_damage_group", to: "right_peptide2", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "right_peptide_to_assays", from: "right_peptide2", to: "right_assay_group", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },
  { id: "right_assays_to_conclusion", from: "right_assay_group", to: "right_conclusion", direction: "forward", route: { kind: "straight", fromSide: "bottom", toSide: "top" }, lineStyle: "solid" },

  { id: "bottom_title_to_model", from: "bottom_title", to: "bottom_model", direction: "forward", route: "right block arrow", lineStyle: "solid", render: "stageArrow", frame: [350, 822, 40, 63] },
  { id: "bottom_model_to_assays", from: "bottom_model", to: "bottom_assay_group", direction: "forward", route: "right block arrow", lineStyle: "solid", render: "stageArrow", frame: [609, 821, 56, 67] },
  { id: "bottom_assays_to_result", from: "bottom_assay_group", to: "bottom_result", direction: "forward", route: "right block arrow", lineStyle: "solid", render: "stageArrow", frame: [978, 814, 59, 81] },
];

// Intrinsic raster insets remain local, replaceable crops of source.png.
const INSETS = [
  { id: "left_qpcr_icon", sourcePixelBox: [58, 275, 69, 66], slideFrame: [58, 275, 69, 66], semanticRole: "RT-qPCR experimental illustration" },
  { id: "left_pipette_icon", sourcePixelBox: [337, 284, 25, 91], slideFrame: [337, 284, 25, 91], semanticRole: "pipette illustration" },
  { id: "left_molecule_icon", sourcePixelBox: [270, 428, 100, 101], slideFrame: [270, 428, 100, 101], semanticRole: "molecular model illustration" },
  { id: "bottom_mouse", sourcePixelBox: [405, 730, 171, 97], slideFrame: [405, 730, 171, 97], semanticRole: "type-2 diabetes animal model mouse" },
  { id: "bottom_msp_icon", sourcePixelBox: [716, 698, 57, 67], slideFrame: [716, 698, 57, 67], semanticRole: "MSP test-tube and pipette illustration" },
  { id: "bottom_chip_icon", sourcePixelBox: [884, 698, 58, 74], slideFrame: [884, 698, 58, 74], semanticRole: "ChIP plate and magnifier illustration" },
  { id: "bottom_rtqpcr_icon", sourcePixelBox: [698, 778, 121, 39], slideFrame: [698, 778, 121, 39], semanticRole: "RT-qPCR device illustration" },
  { id: "bottom_wb_icon", sourcePixelBox: [863, 772, 43, 46], slideFrame: [863, 772, 43, 46], semanticRole: "Western blot molecular illustration" },
  { id: "bottom_elisa_icon", sourcePixelBox: [719, 863, 77, 50], slideFrame: [719, 863, 77, 50], semanticRole: "ELISA illustration" },
  { id: "bottom_if_icon", sourcePixelBox: [845, 881, 86, 31], slideFrame: [845, 881, 86, 31], semanticRole: "immunofluorescence slide illustration" },
  { id: "bottom_iron_swatch", sourcePixelBox: [783, 946, 71, 37], slideFrame: [783, 946, 71, 37], semanticRole: "iron-ion colorimetric swatches" },
];

function position(bounds) {
  const [left, top, width, height] = bounds;
  return { left, top, width, height };
}

function styleFor(key) {
  const common = { geometry: "roundRect", fill: COLORS.lightBlue, line: { style: "solid", fill: COLORS.lightBlueBorder, width: 1.2 }, borderRadius: 8 };
  const styles = {
    title: { geometry: "roundRect", fill: COLORS.title, line: { style: "solid", fill: COLORS.titleBorder, width: 1.2 }, borderRadius: 18 },
    blue: common,
    blueLarge: { ...common, borderRadius: 28 },
    relation: { geometry: "roundRect", fill: "#F7FCFF", line: { style: "solid", fill: COLORS.relationBorder, width: 1.2 }, borderRadius: 8 },
    cream: { geometry: "roundRect", fill: COLORS.cream, line: { style: "solid", fill: COLORS.creamBorder, width: 1.0 }, borderRadius: 8 },
    mint: { geometry: "roundRect", fill: COLORS.mint, line: { style: "solid", fill: COLORS.mintBorder, width: 1.0 }, borderRadius: 6 },
    yellow: { geometry: "roundRect", fill: COLORS.yellow, line: { style: "solid", fill: COLORS.yellowBorder, width: 1.2 }, borderRadius: 16 },
    pale: { geometry: "roundRect", fill: COLORS.paleBlue, line: { style: "solid", fill: "#4E7FA1", width: 1.4 }, borderRadius: 22, shadow: "shadow-sm" },
    paleLeft: { geometry: "roundRect", fill: COLORS.paleBlue, line: { style: "solid", fill: "#4E7FA1", width: 1.4 }, borderRadius: 28, shadow: "shadow-sm" },
    dashedGray: { geometry: "roundRect", fill: "none", line: { style: "dashed", fill: COLORS.dash, width: 1.5 }, borderRadius: 16 },
    dashedBlue: { geometry: "roundRect", fill: "none", line: { style: "dashed", fill: COLORS.dashBlue, width: 1.7 }, borderRadius: 26 },
    dashedBlueLarge: { geometry: "roundRect", fill: "none", line: { style: "dashed", fill: COLORS.dashBlue, width: 2.5 }, borderRadius: 48 },
    textOnly: { geometry: "textbox", fill: "none", line: { style: "solid", fill: "none", width: 0 } },
  };
  return styles[key] ?? common;
}

function richRunsForLine(line) {
  const regex = /(L\. ?plantarum|β)/g;
  const runs = [];
  let cursor = 0;
  for (const match of line.matchAll(regex)) {
    if (match.index > cursor) runs.push(line.slice(cursor, match.index));
    if (match[0] === "β") {
      runs.push({ run: match[0], textStyle: { typeface: LATIN } });
    } else {
      runs.push({ run: match[0], textStyle: { italic: true, typeface: LATIN } });
    }
    cursor = match.index + match[0].length;
  }
  if (cursor < line.length) runs.push(line.slice(cursor));
  return runs;
}

function setNodeText(shape, node) {
  if (!node.text) return;
  if (node.rich) {
    shape.text.set(node.text.split("\n").map(richRunsForLine));
  } else {
    shape.text = node.text;
  }
  shape.text.style = {
    fontSize: node.fontSize ?? 18,
    bold: node.bold ?? node.style === "title",
    color: node.style === "title" ? "#FFFFFF" : "#111111",
    alignment: node.style === "paleLeft" ? "left" : "center",
    verticalAlignment: "middle",
    autoFit: "none",
    wrap: "square",
    typeface: node.typeface ?? SONG,
    lineSpacing: node.style === "title" ? 0.92 : 0.96,
    insets: node.style === "paleLeft"
      ? { top: 7, right: 12, bottom: 7, left: 20 }
      : { top: 2, right: 6, bottom: 2, left: 6 },
  };
}

// Artifact Tool 2.8.39 maps tail to the target/end and head to the source/start.
function connectOneWay(slide, from, to, route = {}) {
  return slide.shapes.connect(from, to, {
    kind: route.kind ?? "elbow",
    fromSide: route.fromSide,
    toSide: route.toSide,
    line: { style: "solid", fill: COLORS.edge, width: 1.25 },
    tail: { type: "triangle", width: "sm", length: "sm" },
  });
}

function connectBothWays(slide, from, to, route = {}) {
  return slide.shapes.connect(from, to, {
    kind: route.kind ?? "elbow",
    fromSide: route.fromSide,
    toSide: route.toSide,
    line: { style: "solid", fill: COLORS.edge, width: 1.25 },
    head: { type: "triangle", width: "sm", length: "sm" },
    tail: { type: "triangle", width: "sm", length: "sm" },
  });
}

function connectExactPath(slide, from, to, route = {}) {
  return connectOneWay(slide, from, to, route);
}

function cropFromPixelBox([x, y, width, height]) {
  return {
    left: x / SOURCE_WIDTH,
    top: y / SOURCE_HEIGHT,
    right: 1 - (x + width) / SOURCE_WIDTH,
    bottom: 1 - (y + height) / SOURCE_HEIGHT,
  };
}

async function assertOutputAbsent() {
  try {
    await fs.access(OUTPUT);
    throw new Error(`Refusing to overwrite existing output: ${OUTPUT}`);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function main() {
  await assertOutputAbsent();
  const sourceBytes = new Uint8Array(await fs.readFile(SOURCE));

  const presentation = Presentation.create({ slideSize: { width: SLIDE_WIDTH, height: SLIDE_HEIGHT } });
  const slide = presentation.slides.add();
  slide.background.fill = COLORS.background;

  slide.shapes.add({
    geometry: "rect",
    name: "source-canvas-frame",
    position: { left: 0.5, top: 0.5, width: SLIDE_WIDTH - 1, height: SLIDE_HEIGHT - 1 },
    fill: "none",
    line: { style: "solid", fill: "#B4B4B4", width: 0.8 },
  });

  // Source section separators.
  for (const [name, bounds] of [
    ["separator-left-center", [410, 5, 1, 657]],
    ["separator-center-right", [901, 5, 1, 657]],
    ["separator-right-summary", [1266, 5, 1, 657]],
    ["separator-top-bottom", [5, 662, 1462, 1]],
  ]) {
    slide.shapes.add({
      geometry: "line",
      name,
      position: position(bounds),
      fill: "none",
      line: { style: "dashed", fill: COLORS.dash, width: 1.25 },
    });
  }

  const shapes = new Map();
  for (const node of NODES) {
    if (node.draw === false) continue;
    const style = styleFor(node.style);
    const shape = slide.shapes.add({
      geometry: style.geometry,
      name: node.id,
      position: position(node.bounds),
      fill: style.fill,
      line: style.line,
      ...(style.borderRadius !== undefined ? { borderRadius: style.borderRadius } : {}),
      ...(style.shadow ? { shadow: style.shadow } : {}),
    });
    setNodeText(shape, node);
    shapes.set(node.id, shape);
  }

  for (const edge of EDGES) {
    if (edge.render === "stageArrow") {
      slide.shapes.add({
        geometry: "rightArrow",
        name: edge.id,
        position: position(edge.frame),
        fill: COLORS.stageArrow,
        line: { style: "solid", fill: COLORS.stageArrowBorder, width: 1.1 },
        shadow: "shadow-sm",
      });
      continue;
    }
    const from = shapes.get(edge.from);
    const to = shapes.get(edge.to);
    if (!from || !to) throw new Error(`Missing endpoint for edge ${edge.id}`);
    if (edge.direction === "bidirectional") {
      connectBothWays(slide, from, to, edge.route);
    } else if (edge.route?.kind?.startsWith("elbow")) {
      connectExactPath(slide, from, to, edge.route);
    } else {
      connectOneWay(slide, from, to, edge.route);
    }
  }

  for (const inset of INSETS) {
    slide.images.add({
      blob: sourceBytes,
      contentType: "image/png",
      alt: inset.semanticRole,
      crop: cropFromPixelBox(inset.sourcePixelBox),
      position: position(inset.slideFrame),
      geometry: "rect",
    });
  }

  slide.speakerNotes.textFrame.setText(
    "[Sources]\n- User-provided source image: source.png (local, delivered beside build.mjs)\n[/Sources]",
  );

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
