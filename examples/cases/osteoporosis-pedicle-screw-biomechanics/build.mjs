/**
 * Native editable reconstruction of the supplied osteoporosis pedicle-screw
 * biomechanics research flowchart.
 *
 * Runtime: @oai/artifact-tool 2.8.39 (Artifact Tool backend).
 * Tested command: node build.mjs
 * The adjacent source.png is read for provenance validation. The slide uses a
 * single source-to-slide transform (1 source pixel = 1 slide pixel), and every
 * meaning-bearing item is rebuilt as native PowerPoint text, shape, line, or
 * connector content. No tracing bitmap is placed on the slide.
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE_PATH = path.join(HERE, "source.png");
const OUTPUT_PATH = path.join(HERE, "editable.pptx");
const SOURCE = { width: 1043, height: 1508 };
const SLIDE = { ...SOURCE };
const SCALE = 1;
const FONT = "Source Han Sans CN";

const C = {
  white: "#FFFFFF",
  canvas: "#FCFCFD",
  ink: "#111111",
  muted: "#4B5563",
  frame: "#171717",
  titleFill: "#ADB9DA",
  titleLine: "#7E91C3",
  tabBlue: "#BCDAF2",
  tabGreen: "#BADFD1",
  tabPink: "#F8D8DA",
  tabOrange: "#FEE7BF",
  headerBlue: "#DEEBFA",
  headerGreen: "#DCEDE8",
  headerPink: "#FCEBEC",
  headerOrange: "#FDF1E5",
  paleBlue: "#E7EFF8",
  paleBlueLine: "#9AAEC8",
  paleGreen: "#D9EEE5",
  paleGreenLine: "#7DB89C",
  palePink: "#FBE7E9",
  palePinkLine: "#E9A9B0",
  palePeach: "#FFF0E3",
  palePeachLine: "#EFC190",
  paleYellow: "#FBEBC0",
  paleYellowLine: "#D5AF51",
  palePurple: "#EEE2F8",
  palePurpleLine: "#A98BCC",
  resultGreen: "#E6F1E3",
  resultGreenLine: "#9CB88F",
  blue: "#2779D1",
  green: "#4C862A",
  red: "#EE9895",
  brown: "#A87345",
  orangeRed: "#F05A21",
};

const STYLES = {
  title: {
    geometry: "roundRect", fill: C.titleFill, line: C.titleLine,
    fontSize: 25.5, bold: true, align: "center", radius: 7,
  },
  section: {
    geometry: "roundRect", fill: C.tabBlue, line: C.tabBlue,
    fontSize: 25, bold: true, align: "center", radius: 7, lineSpacing: 0.88,
  },
  panelTitle: {
    geometry: "rect", fill: C.headerBlue, line: C.frame,
    fontSize: 18.8, bold: true, align: "center", lineWidth: 1.15,
  },
  body: {
    geometry: "textbox", fill: "none", line: "none",
    fontSize: 17, bold: false, align: "left", lineSpacing: 1.18,
  },
  bodyCenter: {
    geometry: "textbox", fill: "none", line: "none",
    fontSize: 17, bold: false, align: "center", lineSpacing: 1.2,
  },
  bodyBold: {
    geometry: "textbox", fill: "none", line: "none",
    fontSize: 17, bold: true, align: "left", lineSpacing: 1.08,
  },
  outlineBox: {
    geometry: "roundRect", fill: C.white, line: C.frame,
    fontSize: 16.5, bold: false, align: "center", radius: 7, lineWidth: 1.25,
  },
  chipBlue: {
    geometry: "roundRect", fill: C.paleBlue, line: C.paleBlueLine,
    fontSize: 15.5, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  chipGreen: {
    geometry: "roundRect", fill: C.paleGreen, line: C.paleGreenLine,
    fontSize: 15.5, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  chipPink: {
    geometry: "roundRect", fill: C.palePink, line: C.palePinkLine,
    fontSize: 15.5, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  chipPeach: {
    geometry: "roundRect", fill: C.palePeach, line: C.palePeachLine,
    fontSize: 15, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  chipYellow: {
    geometry: "roundRect", fill: C.paleYellow, line: C.paleYellowLine,
    fontSize: 15, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  chipPurple: {
    geometry: "roundRect", fill: C.palePurple, line: C.palePurpleLine,
    fontSize: 14.5, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  chipResult: {
    geometry: "roundRect", fill: C.resultGreen, line: C.resultGreenLine,
    fontSize: 15, bold: false, align: "center", radius: 6, lineWidth: 0.85,
  },
  smallLabel: {
    geometry: "textbox", fill: "none", line: "none",
    fontSize: 15.5, bold: true, align: "center", lineSpacing: 0.92,
  },
  index: {
    geometry: "textbox", fill: "none", line: "none",
    fontSize: 18, bold: true, align: "right", lineSpacing: 0.9,
  },
  plus: {
    geometry: "textbox", fill: "none", line: "none",
    fontSize: 20, bold: true, align: "center",
  },
  anchor: {
    geometry: "ellipse", fill: "none", line: "none",
    fontSize: 1, bold: false, align: "center",
  },
};

// Source-pixel semantic object map. Bounds are [x, y, width, height] on the
// exact 1043 x 1508 reference canvas. Vertical labels use one editable text box
// with authored line breaks; glyphs remain upright and are not fragmented.
const NODES = [
  { id: "g_title", text: "探究不同椎弓根螺钉固定策略对骨质疏松性胸腰椎骨折稳定性的生物力学影响", type: "title", bounds: [6, 9, 1002, 59], style: "title" },

  { id: "tab_review", text: "综\n述", type: "section-label", bounds: [6, 79, 53, 286], style: "section", fill: C.tabBlue },
  { id: "tab_parameters", text: "关\n键\n参\n数\n提\n取\n与\n测\n量", type: "section-label", bounds: [6, 376, 53, 413], style: "section", fill: C.tabGreen, fontSize: 23, lineSpacing: 0.82 },
  { id: "tab_methods", text: "方\n法\n整\n合\n与\n验\n证", type: "section-label", bounds: [6, 805, 52, 360], style: "section", fill: C.tabPink, fontSize: 23.5, lineSpacing: 0.82 },
  { id: "tab_mechanism", text: "机\n理\n解\n析\n与\n效\n果\n预\n测", type: "section-label", bounds: [6, 1194, 50, 270], style: "section", fill: C.tabOrange, fontSize: 21.5, lineSpacing: 0.75 },

  { id: "p11_title", text: "骨质疏松性胸腰椎骨折的基本概念与特征", type: "panel-title", bounds: [83, 89, 335, 41], style: "panelTitle", fill: C.headerBlue, fontSize: 17.8 },
  { id: "p11_body", text: "骨质疏松导致骨小梁退化\n胸腰段应力集中，易发生压缩骨折\n椎弓根螺钉固定是主要治疗策略", type: "body", bounds: [95, 140, 264, 86], style: "body", fontSize: 16.8, lineSpacing: 1.26 },
  { id: "p11_core", text: "核心：减轻疼痛、\n恢复椎体高度", type: "conclusion", bounds: [94, 243, 139, 81], style: "outlineBox", fontSize: 16.2, bold: true },
  { id: "p11_emphasis", text: "强调：提升内固定\n稳定性，降低\n松动风险", type: "conclusion", bounds: [253, 243, 151, 81], style: "outlineBox", fontSize: 15.8, bold: true },
  { id: "p11_lead_label", text: "引出", type: "edge-label", bounds: [371, 153, 38, 26], style: "smallLabel", fontSize: 15 },
  { id: "p11_summary_label", text: "总结", type: "edge-label", bounds: [371, 201, 38, 26], style: "smallLabel", fontSize: 15 },
  { id: "p11_index", text: "1.1", type: "index", bounds: [368, 327, 40, 24], style: "index", color: C.blue },

  { id: "p12_title", text: "当前内固定技术在综合治疗中的应用与进展", type: "panel-title", bounds: [426, 89, 375, 41], style: "panelTitle", fill: C.headerBlue, fontSize: 18 },
  { id: "p12_intro", text: "短节段与长节段固定的适用范围\n导航与机器人辅助置钉技术", type: "body", bounds: [443, 140, 300, 59], style: "body", fontSize: 16.8, lineSpacing: 1.18 },
  { id: "p12_include", text: "包括", type: "edge-label", bounds: [427, 249, 40, 28], style: "smallLabel", fontSize: 15 },
  { id: "p12_stack_1", text: "体内固定松动与拔出机制", type: "stack-row", bounds: [470, 208, 242, 25], style: "bodyCenter", fontSize: 15.3 },
  { id: "p12_stack_2", text: "单皮质固定研究", type: "stack-row", bounds: [470, 234, 242, 26], style: "bodyCenter", fontSize: 15.3 },
  { id: "p12_stack_3", text: "双皮质固定研究", type: "stack-row", bounds: [470, 261, 242, 26], style: "bodyCenter", fontSize: 15.3 },
  { id: "p12_stack_4", text: "有限元研究", type: "stack-row", bounds: [470, 288, 242, 26], style: "bodyCenter", fontSize: 15.3 },
  { id: "p12_stack_5", text: "尸体与模型实验研究", type: "stack-row", bounds: [470, 315, 242, 26], style: "bodyCenter", fontSize: 15.3 },
  { id: "p12_parallel", text: "并列", type: "edge-label", bounds: [733, 157, 42, 26], style: "smallLabel", fontSize: 15 },
  { id: "p12_research", text: "研\n究\n手\n段", type: "edge-label", bounds: [755, 216, 37, 92], style: "smallLabel", fontSize: 14.5, lineSpacing: 0.8 },
  { id: "p12_index", text: "1.2", type: "index", bounds: [751, 327, 40, 24], style: "index", color: C.blue },
  { id: "p12_method_anchor", text: "", type: "anchor", bounds: [779, 272, 2, 2], style: "anchor" },
  { id: "p13_method_anchor", text: "", type: "anchor", bounds: [815, 272, 2, 2], style: "anchor" },

  { id: "p13_title", text: "研究设计和方法", type: "panel-title", bounds: [808, 89, 186, 41], style: "panelTitle", fill: C.headerBlue, fontSize: 18 },
  { id: "p13_methods", text: "CT三维建模\n\n有限元仿真\n\n生物力学实验", type: "body", bounds: [825, 161, 153, 145], style: "bodyCenter", fontSize: 17.2, lineSpacing: 1.0 },
  { id: "p13_index", text: "1.3", type: "index", bounds: [948, 327, 38, 24], style: "index", color: C.blue },

  { id: "p21_title", text: "螺钉特征在形成和治疗中的应用", type: "panel-title", bounds: [83, 387, 330, 43], style: "panelTitle", fill: C.headerGreen, fontSize: 18 },
  { id: "p21_pre_1", text: "术前", type: "phase", bounds: [91, 448, 69, 34], style: "chipGreen", fontSize: 17, bold: true },
  { id: "p21_imaging", text: "骨折影像学特点研究", type: "study", bounds: [158, 448, 186, 34], style: "outlineBox", fontSize: 16.3, bold: true, radius: 5 },
  { id: "p21_ct", text: "60例患者术前CT资料", type: "dataset", bounds: [136, 497, 201, 33], style: "outlineBox", fontSize: 16.1, radius: 4 },
  { id: "p21_morph", text: "骨折节段形态与塌陷程度分析", type: "result", bounds: [143, 561, 204, 35], style: "outlineBox", fontSize: 13.6, insetX: 1, radius: 4 },
  { id: "p21_got_1", text: "得\n到", type: "edge-label", bounds: [96, 538, 27, 48], style: "smallLabel", fontSize: 14, lineSpacing: 0.83 },
  { id: "p21_pre_2", text: "术前", type: "phase", bounds: [91, 611, 69, 34], style: "chipGreen", fontSize: 17, bold: true },
  { id: "p21_mechanism", text: "螺钉固定机制研究", type: "study", bounds: [158, 611, 186, 34], style: "outlineBox", fontSize: 16.3, bold: true, radius: 5 },
  { id: "p21_density", text: "60例患者骨密度与椎弓根参数", type: "dataset", bounds: [136, 660, 202, 34], style: "outlineBox", fontSize: 13.2, insetX: 1, radius: 4 },
  { id: "p21_risk", text: "螺钉松动风险的定量比较", type: "result", bounds: [149, 721, 198, 35], style: "outlineBox", fontSize: 14.2, insetX: 1, radius: 4 },
  { id: "p21_got_2", text: "得\n到", type: "edge-label", bounds: [96, 699, 27, 48], style: "smallLabel", fontSize: 14, lineSpacing: 0.83 },
  { id: "p21_index", text: "2.1", type: "index", bounds: [287, 748, 49, 26], style: "index", color: C.green },

  { id: "p22_title", text: "对特定骨折病例的螺钉参数测量分析", type: "panel-title", bounds: [423, 387, 569, 43], style: "panelTitle", fill: C.headerGreen, fontSize: 18.3 },
  { id: "p22_r1", text: "选择特征参数测量方法", type: "row-label", bounds: [437, 444, 208, 31], style: "bodyBold", fontSize: 16.2 },
  { id: "p22_r1_2d", text: "二维", type: "parameter", bounds: [650, 444, 58, 33], style: "chipBlue", fontSize: 16 },
  { id: "p22_r1_3d", text: "三维", type: "parameter", bounds: [717, 444, 58, 33], style: "chipGreen", fontSize: 16 },
  { id: "p22_r2", text: "骨质疏松合并胸腰椎骨折患者的特殊情况", type: "row-label", bounds: [437, 493, 290, 31], style: "bodyBold", fontSize: 15.3 },
  { id: "p22_r2_risk", text: "第一次骨折—再次骨折风险", type: "parameter", bounds: [733, 488, 202, 38], style: "chipGreen", fontSize: 15.3 },
  { id: "p22_r3", text: "定义关键参数及其测量技术", type: "row-label", bounds: [437, 539, 227, 34], style: "bodyBold", fontSize: 16.1 },
  { id: "p22_r3_seq", text: "序列影像参数", type: "parameter", bounds: [671, 539, 120, 34], style: "chipBlue", fontSize: 15.2 },
  { id: "p22_r3_bone", text: "骨质条件参数", type: "parameter", bounds: [798, 539, 137, 34], style: "chipGreen", fontSize: 15.2 },
  { id: "p22_r4", text: "X线影像的二维参数测量结果", type: "row-label", bounds: [437, 590, 238, 44], style: "bodyBold", fontSize: 15.8 },
  { id: "p22_r4_values", text: "椎体前缘高度、后缘高度、椎弓根宽度、\nCobb角、椎体楔角", type: "parameter", bounds: [681, 585, 290, 49], style: "chipGreen", fontSize: 14.5, align: "left" },
  { id: "p22_r5", text: "基于三维CT的螺钉重建", type: "row-label", bounds: [437, 643, 215, 43], style: "bodyBold", fontSize: 16 },
  { id: "p22_r5_traj", text: "轨迹", type: "parameter", bounds: [658, 646, 60, 34], style: "chipGreen", fontSize: 15.5 },
  { id: "p22_r5_density", text: "骨密度", type: "parameter", bounds: [728, 646, 70, 34], style: "chipGreen", fontSize: 15.3 },
  { id: "p22_r5_distance", text: "有效钉道与皮质距离", type: "parameter", bounds: [807, 646, 164, 34], style: "chipGreen", fontSize: 14.5 },
  { id: "p22_r6", text: "重建模型的三维参数测量结果", type: "row-label", bounds: [437, 694, 232, 39], style: "bodyBold", fontSize: 15.8 },
  { id: "p22_r6_corr", text: "序列长度与骨密度的相关分析", type: "parameter", bounds: [673, 694, 247, 35], style: "chipGreen", fontSize: 15.1 },
  { id: "p22_conclusion", text: "完善测量参数体系与提升测量准确度", type: "conclusion", bounds: [548, 746, 318, 29], style: "bodyBold", fontSize: 15.2, align: "center" },
  { id: "p22_index", text: "2.2", type: "index", bounds: [941, 748, 43, 26], style: "index", color: C.green },
  { id: "bus_special_target", text: "", type: "anchor", bounds: [420, 506, 2, 2], style: "anchor" },
  { id: "bus_r5_target", text: "", type: "anchor", bounds: [420, 653, 2, 2], style: "anchor" },
  { id: "bus_conclusion_target", text: "", type: "anchor", bounds: [492, 760, 2, 2], style: "anchor" },
  { id: "bus_green_start_1", text: "", type: "anchor", bounds: [389, 810, 2, 2], style: "anchor" },
  { id: "bus_green_start_2", text: "", type: "anchor", bounds: [399, 810, 2, 2], style: "anchor" },

  { id: "s3_model_integration", text: "模\n型\n整\n合", type: "edge-label", bounds: [78, 852, 41, 67], style: "smallLabel", fontSize: 15.2, lineSpacing: 0.82 },
  { id: "s3_validation", text: "验\n证\n方\n法", type: "edge-label", bounds: [78, 939, 41, 67], style: "smallLabel", fontSize: 15.2, lineSpacing: 0.82 },
  { id: "s3_refinement", text: "模\n型\n细\n分", type: "edge-label", bounds: [78, 1025, 41, 67], style: "smallLabel", fontSize: 15.2, lineSpacing: 0.82 },
  { id: "red_bus_bottom", text: "", type: "anchor", bounds: [367, 1202, 2, 2], style: "anchor" },
  { id: "p42_bus_target", text: "", type: "anchor", bounds: [396, 1202, 2, 2], style: "anchor" },

  { id: "p31_title", text: "有限元模型建立", type: "panel-title", bounds: [148, 819, 186, 41], style: "panelTitle", fill: C.headerPink, fontSize: 18.2 },
  { id: "p31_body", text: "重建胸腰椎骨折三维模型\n网格划分和材料赋值\n韧带与界面建模\n有效性验证（与实验对比）", type: "body", bounds: [156, 868, 171, 132], style: "body", fontSize: 13.2, insetX: 0.5, lineSpacing: 1.4 },
  { id: "p31_summary", text: "模型综合分析：\n测量方法的可视化应用；\n发展为后续仿真平台", type: "result", bounds: [157, 1021, 155, 112], style: "outlineBox", fill: C.palePink, fontSize: 12.4, insetX: 1, lineSpacing: 1.3, radius: 9 },
  { id: "p31_index", text: "3.1", type: "index", bounds: [289, 1136, 38, 23], style: "index", color: C.orangeRed },

  { id: "p32_title", text: "单皮质固定（仿真）", type: "panel-title", bounds: [346, 819, 208, 41], style: "panelTitle", fill: C.headerPink, fontSize: 17.8 },
  { id: "p32_position", text: "进钉位置", type: "factor", bounds: [357, 868, 72, 29], style: "chipPeach", fontSize: 14.6 },
  { id: "p32_angle", text: "角度变化", type: "factor", bounds: [449, 868, 87, 29], style: "chipPeach", fontSize: 14.6 },
  { id: "p32_single", text: "单皮质", type: "fixation", bounds: [400, 909, 96, 28], style: "chipPink", fontSize: 16 },
  { id: "p32_len", text: "不同钉长", type: "parameter", bounds: [350, 947, 63, 31], style: "chipPeach", fontSize: 11.3, insetX: 0.5 },
  { id: "p32_diam", text: "不同钉径", type: "parameter", bounds: [415, 947, 63, 31], style: "chipPeach", fontSize: 11.3, insetX: 0.5 },
  { id: "p32_depth", text: "不同进钉深度", type: "parameter", bounds: [480, 947, 70, 31], style: "chipPeach", fontSize: 10.1, insetX: 0.3 },
  { id: "p32_loading", text: "活加载", type: "edge-label", bounds: [328, 986, 38, 24], style: "smallLabel", fontSize: 10.8, insetX: 0 },
  { id: "p32_load_axial", text: "轴向压缩", type: "load", bounds: [354, 1002, 43, 34], style: "chipPeach", fontSize: 9.5, insetX: 0.2 },
  { id: "p32_load_flex", text: "前屈", type: "load", bounds: [399, 1002, 34, 34], style: "chipPeach", fontSize: 11.1, insetX: 0.2 },
  { id: "p32_load_ext", text: "后伸", type: "load", bounds: [435, 1002, 34, 34], style: "chipPeach", fontSize: 11.1, insetX: 0.2 },
  { id: "p32_load_lat", text: "侧弯", type: "load", bounds: [471, 1002, 34, 34], style: "chipPeach", fontSize: 11.1, insetX: 0.2 },
  { id: "p32_load_rot", text: "轴向旋转", type: "load", bounds: [507, 1002, 43, 34], style: "chipPeach", fontSize: 9.3, insetX: 0.2 },
  { id: "p32_outcome", text: "评估单皮质固定的\n应力分布与稳定性", type: "outcome", bounds: [362, 1073, 176, 62], style: "outlineBox", fill: C.palePink, fontSize: 15.3, lineSpacing: 1.04, radius: 8 },
  { id: "p32_index", text: "3.2", type: "index", bounds: [509, 1136, 38, 23], style: "index", color: C.orangeRed },

  { id: "p33_title", text: "双皮质固定（仿真）", type: "panel-title", bounds: [559, 819, 220, 41], style: "panelTitle", fill: C.headerPink, fontSize: 17.8 },
  { id: "p33_penetration", text: "穿透长度", type: "factor", bounds: [618, 868, 96, 29], style: "chipPeach", fontSize: 14.6 },
  { id: "p33_double", text: "双皮质", type: "fixation", bounds: [620, 909, 96, 28], style: "chipPink", fontSize: 16 },
  { id: "p33_compare", text: "与单皮质比较", type: "parameter", bounds: [562, 947, 69, 31], style: "chipPeach", fontSize: 10.5, insetX: 0.3 },
  { id: "p33_pen_len", text: "不同穿透长度", type: "parameter", bounds: [633, 947, 71, 31], style: "chipPeach", fontSize: 9.9, insetX: 0.3 },
  { id: "p33_traj", text: "不同轨迹角度", type: "parameter", bounds: [706, 947, 70, 31], style: "chipPeach", fontSize: 9.9, insetX: 0.3 },
  { id: "p33_loading", text: "活加载", type: "edge-label", bounds: [540, 986, 38, 24], style: "smallLabel", fontSize: 10.8, insetX: 0 },
  { id: "p33_load_axial", text: "轴向压缩", type: "load", bounds: [566, 1002, 43, 34], style: "chipPeach", fontSize: 9.5, insetX: 0.2 },
  { id: "p33_load_flex", text: "前屈", type: "load", bounds: [611, 1002, 34, 34], style: "chipPeach", fontSize: 11.1, insetX: 0.2 },
  { id: "p33_load_ext", text: "后伸", type: "load", bounds: [647, 1002, 34, 34], style: "chipPeach", fontSize: 11.1, insetX: 0.2 },
  { id: "p33_load_lat", text: "侧弯", type: "load", bounds: [683, 1002, 34, 34], style: "chipPeach", fontSize: 11.1, insetX: 0.2 },
  { id: "p33_load_rot", text: "轴向旋转", type: "load", bounds: [719, 1002, 57, 34], style: "chipPeach", fontSize: 9.5, insetX: 0.2 },
  { id: "p33_outcome", text: "验证双皮质固定的\n抗拔出能力与抗失稳优势", type: "outcome", bounds: [575, 1073, 188, 62], style: "outlineBox", fill: C.palePink, fontSize: 14.5, lineSpacing: 1.04, radius: 8 },
  { id: "p33_index", text: "3.3", type: "index", bounds: [734, 1136, 38, 23], style: "index", color: C.orangeRed },

  { id: "p34_title", text: "双皮质固定（实验）", type: "panel-title", bounds: [784, 819, 208, 41], style: "panelTitle", fill: C.headerPink, fontSize: 17.8 },
  { id: "p34_collect", text: "样本采集、标准化固定", type: "step", bounds: [801, 868, 174, 33], style: "chipPeach", fontSize: 14.8 },
  { id: "p34_specimen", text: "骨质疏松椎体替代体样本", type: "step", bounds: [795, 943, 181, 36], style: "chipPeach", fontSize: 14.1 },
  { id: "p34_torque", text: "力矩与活动度", type: "result", bounds: [797, 1016, 111, 38], style: "chipPeach", fontSize: 14.7 },
  { id: "p34_stiffness", text: "刚度", type: "result", bounds: [930, 1016, 61, 38], style: "chipPeach", fontSize: 15.5 },
  { id: "p34_outcome", text: "固定过程中的操作难点\n与安全边界", type: "outcome", bounds: [798, 1073, 185, 62], style: "outlineBox", fill: C.palePink, fontSize: 15.2, lineSpacing: 1.04, radius: 8 },
  { id: "p34_index", text: "3.4", type: "index", bounds: [947, 1136, 38, 23], style: "index", color: C.orangeRed },

  { id: "p41_title", text: "螺钉固定的稳定机制解析", type: "panel-title", bounds: [80, 1201, 263, 40], style: "panelTitle", fill: C.headerOrange, fontSize: 17.8 },
  { id: "p41_interface", text: "骨—钉界面模型", type: "model", bounds: [138, 1253, 147, 33], style: "chipBlue", fontSize: 16 },
  { id: "p41_stress", text: "应力", type: "metric", bounds: [88, 1309, 52, 35], style: "chipPeach", fontSize: 15.5 },
  { id: "p41_stability", text: "稳定性", type: "metric", bounds: [160, 1305, 83, 43], style: "chipYellow", fontSize: 16, bold: true },
  { id: "p41_displacement", text: "位移", type: "metric", bounds: [270, 1309, 52, 35], style: "chipPink", fontSize: 15.5 },
  { id: "p41_extract", text: "应力提取技术", type: "output", bounds: [91, 1385, 103, 37], style: "chipResult", fontSize: 14.8 },
  { id: "p41_optimize", text: "固定策略优化", type: "output", bounds: [216, 1385, 105, 37], style: "chipResult", fontSize: 14.8 },
  { id: "p41_index", text: "4.1", type: "index", bounds: [298, 1425, 37, 23], style: "index", color: C.orangeRed },

  { id: "p42_title", text: "单皮质固定效果预测", type: "panel-title", bounds: [402, 1203, 259, 38], style: "panelTitle", fill: C.headerOrange, fontSize: 17.5 },
  { id: "p42_stress", text: "骨质应力", type: "input", bounds: [407, 1252, 67, 30], style: "chipYellow", fontSize: 14.1 },
  { id: "p42_plus_1", text: "+", type: "operator", bounds: [477, 1253, 18, 28], style: "plus" },
  { id: "p42_length", text: "钉长变化", type: "input", bounds: [496, 1252, 70, 30], style: "chipYellow", fontSize: 14.1 },
  { id: "p42_plus_2", text: "+", type: "operator", bounds: [568, 1253, 18, 28], style: "plus" },
  { id: "p42_diameter", text: "钉径变化", type: "input", bounds: [588, 1252, 67, 30], style: "chipYellow", fontSize: 14.1 },
  { id: "p42_formula", text: "公式", type: "formula", bounds: [436, 1296, 45, 31], style: "outlineBox", fontSize: 14, radius: 4 },
  { id: "p42_prediction", text: "预测稳定性响应幅度", type: "prediction", bounds: [508, 1295, 137, 33], style: "chipPurple", fontSize: 13.7 },
  { id: "p42_index", text: "4.2", type: "index", bounds: [404, 1315, 38, 22], style: "index", color: C.orangeRed, align: "left" },

  { id: "p43_title", text: "双皮质固定效果预测", type: "panel-title", bounds: [682, 1203, 301, 38], style: "panelTitle", fill: C.headerOrange, fontSize: 17.5 },
  { id: "p43_stress", text: "皮质应力", type: "input", bounds: [688, 1252, 68, 30], style: "chipYellow", fontSize: 14.1 },
  { id: "p43_plus_1", text: "+", type: "operator", bounds: [759, 1253, 18, 28], style: "plus" },
  { id: "p43_length", text: "穿透长度", type: "input", bounds: [779, 1252, 75, 30], style: "chipYellow", fontSize: 14.1 },
  { id: "p43_plus_2", text: "+", type: "operator", bounds: [858, 1253, 18, 28], style: "plus" },
  { id: "p43_angle", text: "轨迹角度", type: "input", bounds: [879, 1252, 76, 30], style: "chipYellow", fontSize: 14.1 },
  { id: "p43_formula", text: "公式", type: "formula", bounds: [719, 1296, 45, 31], style: "outlineBox", fontSize: 14, radius: 4 },
  { id: "p43_prediction", text: "预测拔出强度响应幅度", type: "prediction", bounds: [791, 1295, 178, 33], style: "chipPurple", fontSize: 13.5 },
  { id: "p43_index", text: "4.3", type: "index", bounds: [684, 1315, 38, 22], style: "index", color: C.orangeRed, align: "left" },

  { id: "p44_title", text: "椎弓根螺钉固定效果预测规则分析", type: "panel-title", bounds: [402, 1347, 582, 35], style: "panelTitle", fill: C.headerOrange, fontSize: 17.7 },
  { id: "p44_index", text: "4.4", type: "index", bounds: [941, 1354, 36, 22], style: "index", color: C.orangeRed },
  { id: "p44_vertebra", text: "椎体", type: "input", bounds: [419, 1395, 72, 34], style: "chipYellow", fontSize: 15.4 },
  { id: "p44_plus", text: "+", type: "operator", bounds: [500, 1398, 18, 28], style: "plus" },
  { id: "p44_pedicle", text: "椎弓根", type: "input", bounds: [530, 1395, 85, 34], style: "chipYellow", fontSize: 15.4 },
  { id: "p44_global", text: "全局参数数据", type: "step", bounds: [678, 1395, 108, 34], style: "chipYellow", fontSize: 14.8 },
  { id: "p44_threshold", text: "全局失效阀值判定", type: "step", bounds: [845, 1395, 128, 34], style: "chipYellow", fontSize: 14.1 },

  { id: "blue_bus_top", text: "", type: "anchor", bounds: [1023, 209, 2, 2], style: "anchor" },
  { id: "blue_bus_bottom", text: "", type: "anchor", bounds: [1023, 570, 2, 2], style: "anchor" },
  { id: "blue_target_top", text: "", type: "anchor", bounds: [992, 209, 2, 2], style: "anchor" },
  { id: "blue_target_bottom", text: "", type: "anchor", bounds: [1001, 570, 2, 2], style: "anchor" },
  { id: "blue_bus_label", text: "方\n法\n对\n应", type: "edge-label", bounds: [1006, 341, 31, 94], style: "smallLabel", fontSize: 14.5, lineSpacing: 0.82 },
  { id: "red_bus_top", text: "", type: "anchor", bounds: [1036, 232, 2, 2], style: "anchor" },
  { id: "red_bus_lower", text: "", type: "anchor", bounds: [1036, 970, 2, 2], style: "anchor" },
  { id: "red_target_top", text: "", type: "anchor", bounds: [998, 232, 2, 2], style: "anchor" },
  { id: "red_target_lower", text: "", type: "anchor", bounds: [994, 970, 2, 2], style: "anchor" },
  { id: "red_bus_label", text: "方\n法\n对\n应", type: "edge-label", bounds: [1008, 674, 30, 98], style: "smallLabel", fontSize: 14.5, lineSpacing: 0.82 },
  { id: "green_bus_top", text: "", type: "anchor", bounds: [1016, 720, 2, 2], style: "anchor" },
  { id: "green_bus_bottom", text: "", type: "anchor", bounds: [1022, 1342, 2, 2], style: "anchor" },
  { id: "green_target_top", text: "", type: "anchor", bounds: [991, 720, 2, 2], style: "anchor" },
  { id: "green_target_bottom", text: "", type: "anchor", bounds: [1000, 1342, 2, 2], style: "anchor" },
  { id: "green_bus_label", text: "内\n容\n输\n出", type: "edge-label", bounds: [1007, 1149, 31, 104], style: "smallLabel", fontSize: 14.5, lineSpacing: 0.82 },
];

const FRAMES = [
  { id: "section_1_frame", bounds: [73, 81, 932, 288], role: "section" },
  { id: "section_2_frame", bounds: [73, 376, 932, 419], role: "section" },
  { id: "section_3_frame", bounds: [72, 801, 933, 378], role: "section" },
  { id: "section_4_frame", bounds: [72, 1190, 933, 277], role: "section" },
  { id: "p11_frame", bounds: [83, 89, 335, 267], role: "panel" },
  { id: "p12_frame", bounds: [426, 89, 375, 267], role: "panel" },
  { id: "p13_frame", bounds: [808, 89, 186, 267], role: "panel" },
  { id: "p21_frame", bounds: [83, 387, 330, 395], role: "panel" },
  { id: "p22_frame", bounds: [423, 387, 569, 394], role: "panel" },
  { id: "p31_frame", bounds: [148, 819, 186, 344], role: "panel" },
  { id: "p32_frame", bounds: [346, 819, 208, 344], role: "panel" },
  { id: "p33_frame", bounds: [559, 819, 220, 344], role: "panel" },
  { id: "p34_frame", bounds: [784, 819, 208, 344], role: "panel" },
  { id: "p41_frame", bounds: [80, 1201, 263, 253], role: "panel" },
  { id: "p42_frame", bounds: [402, 1203, 259, 139], role: "panel" },
  { id: "p43_frame", bounds: [682, 1203, 301, 140], role: "panel" },
  { id: "p44_frame", bounds: [402, 1347, 582, 107], role: "panel" },
  { id: "p12_stack_frame", bounds: [470, 206, 242, 136], role: "stack", fill: C.paleBlue },
];

const LINES = [
  { id: "p12_div_1", from: [470, 233], to: [712, 233] },
  { id: "p12_div_2", from: [470, 260], to: [712, 260] },
  { id: "p12_div_3", from: [470, 287], to: [712, 287] },
  { id: "p12_div_4", from: [470, 314], to: [712, 314] },
  { id: "p12_bracket_v", from: [751, 195], to: [751, 302] },
  { id: "p12_bracket_t", from: [745, 195], to: [757, 195] },
  { id: "p12_bracket_b", from: [745, 302], to: [757, 302] },

  { id: "p22_div_1", from: [437, 481], to: [650, 481], color: "#9B9B9B", width: 0.8 },
  { id: "p22_div_2", from: [437, 527], to: [650, 527], color: "#9B9B9B", width: 0.8 },
  { id: "p22_div_3", from: [437, 579], to: [650, 579], color: "#9B9B9B", width: 0.8 },
  { id: "p22_div_4", from: [437, 637], to: [650, 637], color: "#9B9B9B", width: 0.8 },
  { id: "p22_div_5", from: [437, 687], to: [650, 687], color: "#9B9B9B", width: 0.8 },
  { id: "p22_div_6", from: [437, 739], to: [862, 739], color: C.frame, width: 1.2 },
  { id: "p22_conclusion_rule_l", from: [424, 761], to: [489, 761], color: C.green, width: 1.2 },
  { id: "p22_conclusion_rule_r", from: [868, 761], to: [920, 761], color: C.frame, width: 1.1 },

  { id: "green_bus_a", from: [347, 574], to: [369, 574], color: C.green, width: 1.5 },
  { id: "green_bus_b", from: [338, 675], to: [369, 675], color: C.green, width: 1.5 },
  { id: "green_bus_c", from: [369, 574], to: [369, 675], color: C.green, width: 1.5 },
  { id: "green_bus_d", from: [369, 574], to: [391, 574], color: C.green, width: 1.5 },
  { id: "green_bus_e", from: [391, 507], to: [391, 574], color: C.green, width: 1.5 },

  { id: "red_bus_top_h", from: [127, 812], to: [389, 812], color: C.red, width: 1.35 },
  { id: "red_bus_left_v", from: [127, 812], to: [127, 1172], color: C.red, width: 1.35 },
  { id: "red_bus_bottom_h", from: [127, 1172], to: [368, 1172], color: C.red, width: 1.35 },
  { id: "red_bus_drop", from: [368, 1172], to: [368, 1203], color: C.red, width: 1.35 },

  { id: "p32_top_stem", from: [449, 897], to: [449, 909] },
  { id: "p32_param_stem", from: [448, 937], to: [448, 946] },
  { id: "p32_param_bus", from: [381, 946], to: [515, 946] },
  { id: "p32_param_l", from: [381, 946], to: [381, 947] },
  { id: "p32_param_m", from: [447, 946], to: [447, 947] },
  { id: "p32_param_r", from: [515, 946], to: [515, 947] },
  { id: "p32_load_stem", from: [448, 978], to: [448, 993] },
  { id: "p32_load_bus", from: [372, 993], to: [529, 993] },
  { id: "p32_load_1", from: [372, 993], to: [372, 1002] },
  { id: "p32_load_2", from: [413, 993], to: [413, 1002] },
  { id: "p32_load_3", from: [450, 993], to: [450, 1002] },
  { id: "p32_load_4", from: [487, 993], to: [487, 1002] },
  { id: "p32_load_5", from: [529, 993], to: [529, 1002] },
  { id: "p32_outcome_stem", from: [449, 1036], to: [449, 1073] },

  { id: "p33_top_stem", from: [668, 897], to: [668, 909] },
  { id: "p33_param_stem", from: [668, 937], to: [668, 946] },
  { id: "p33_param_bus", from: [596, 946], to: [741, 946] },
  { id: "p33_param_l", from: [596, 946], to: [596, 947] },
  { id: "p33_param_m", from: [668, 946], to: [668, 947] },
  { id: "p33_param_r", from: [741, 946], to: [741, 947] },
  { id: "p33_load_stem", from: [668, 978], to: [668, 993] },
  { id: "p33_load_bus", from: [584, 993], to: [747, 993] },
  { id: "p33_load_1", from: [584, 993], to: [584, 1002] },
  { id: "p33_load_2", from: [625, 993], to: [625, 1002] },
  { id: "p33_load_3", from: [662, 993], to: [662, 1002] },
  { id: "p33_load_4", from: [699, 993], to: [699, 1002] },
  { id: "p33_load_5", from: [747, 993], to: [747, 1002] },
  { id: "p33_outcome_stem", from: [668, 1036], to: [668, 1073] },

  { id: "p34_branch_stem", from: [886, 979], to: [886, 995] },
  { id: "p34_branch_bus", from: [852, 995], to: [960, 995] },
  { id: "p34_branch_l", from: [852, 995], to: [852, 1016] },
  { id: "p34_branch_r", from: [960, 995], to: [960, 1016] },
  { id: "p34_outcome_stem", from: [886, 995], to: [886, 1073] },

  { id: "p41_metric_bus", from: [114, 1296], to: [296, 1296] },
  { id: "p41_metric_l", from: [114, 1296], to: [114, 1309] },
  { id: "p41_metric_m", from: [201, 1286], to: [201, 1305] },
  { id: "p41_metric_r", from: [296, 1296], to: [296, 1309] },
  { id: "p41_bottom_l", from: [114, 1344], to: [114, 1367] },
  { id: "p41_bottom_r", from: [296, 1344], to: [296, 1367] },
  { id: "p41_bottom_bus", from: [114, 1367], to: [296, 1367] },
  { id: "p41_output_l", from: [143, 1367], to: [143, 1385] },
  { id: "p41_output_r", from: [268, 1367], to: [268, 1385] },

  { id: "blue_bus_v1", from: [1024, 210], to: [1024, 339], color: C.blue, width: 1.35 },
  { id: "blue_bus_v2", from: [1024, 435], to: [1024, 571], color: C.blue, width: 1.35 },
  { id: "red_ext_v1", from: [1037, 233], to: [1037, 672], color: C.red, width: 1.15 },
  { id: "red_ext_v2", from: [1037, 773], to: [1037, 971], color: C.red, width: 1.15 },
  { id: "green_ext_v1", from: [1017, 721], to: [1017, 1146], color: C.green, width: 1.25 },
  { id: "green_ext_v2", from: [1023, 1255], to: [1023, 1343], color: C.green, width: 1.25 },
];

// Explicit relationship map. Artifact Tool currently maps `tail` to the target
// end and `head` to the source end; the helper functions below make the
// convention explicit. Block arrows remain one native arrow shape each.
const EDGES = [
  { id: "e_p12_methods_p13", from: "p12_method_anchor", to: "p13_method_anchor", direction: "forward", route: "straight", color: C.blue, width: 1.55 },
  { id: "e_green_bus_special", from: "bus_green_start_1", to: "bus_special_target", direction: "forward", route: "elbow4", fromSide: "top", toSide: "left", color: C.green, width: 1.5 },
  { id: "e_green_bus_r5", from: "bus_green_start_1", to: "bus_r5_target", direction: "forward", route: "elbow3", fromSide: "top", toSide: "left", color: C.green, width: 1.5 },
  { id: "e_green_bus_conclusion", from: "bus_green_start_2", to: "bus_conclusion_target", direction: "forward", route: "elbow3", fromSide: "top", toSide: "left", color: C.green, width: 1.5 },
  { id: "e_red_bus_to_p42", from: "red_bus_bottom", to: "p42_bus_target", direction: "forward", route: "straight", color: C.red, width: 1.3 },

  { id: "e_blue_ext_top", from: "blue_bus_top", to: "blue_target_top", direction: "forward", route: "straight", color: C.blue, width: 1.35 },
  { id: "e_blue_ext_bottom", from: "blue_bus_bottom", to: "blue_target_bottom", direction: "forward", route: "straight", color: C.blue, width: 1.35 },
  { id: "e_red_ext_top", from: "red_bus_top", to: "red_target_top", direction: "forward", route: "straight", color: C.red, width: 1.15 },
  { id: "e_red_ext_bottom", from: "red_bus_lower", to: "red_target_lower", direction: "forward", route: "straight", color: C.red, width: 1.15 },
  { id: "e_green_ext_top", from: "green_bus_top", to: "green_target_top", direction: "forward", route: "straight", color: C.green, width: 1.25 },
  { id: "e_green_ext_bottom", from: "green_bus_bottom", to: "green_target_bottom", direction: "forward", route: "straight", color: C.green, width: 1.25 },

  { id: "e_p21_ct_result", from: "p21_ct", to: "p21_morph", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", color: C.green, width: 1.5 },
  { id: "e_p21_density_result", from: "p21_density", to: "p21_risk", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", color: C.green, width: 1.5 },
  { id: "e_p31_model", from: "s3_model_integration", to: "p31_body", direction: "forward", route: "horizontal-block", visual: "block", geometry: "rightArrow", bounds: [117, 873, 31, 23], color: C.brown },
  { id: "e_p31_validation", from: "s3_validation", to: "p31_body", direction: "forward", route: "horizontal-block", visual: "block", geometry: "rightArrow", bounds: [117, 960, 31, 23], color: C.brown },
  { id: "e_p31_refinement", from: "s3_refinement", to: "p31_summary", direction: "forward", route: "horizontal-block", visual: "block", geometry: "rightArrow", bounds: [117, 1047, 31, 23], color: C.brown },
  { id: "e_p34_collect_specimen", from: "p34_collect", to: "p34_specimen", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", color: C.ink, width: 1.2 },
  { id: "e_p41_up", from: "p41_stability", to: "p41_interface", direction: "forward", route: "straight", fromSide: "top", toSide: "bottom", color: C.ink, width: 1.2 },
  { id: "e_p42_formula", from: "p42_formula", to: "p42_prediction", direction: "forward", route: "straight", fromSide: "right", toSide: "left", color: C.ink, width: 1.2 },
  { id: "e_p43_formula", from: "p43_formula", to: "p43_prediction", direction: "forward", route: "straight", fromSide: "right", toSide: "left", color: C.ink, width: 1.2 },
  { id: "e_p44_global", from: "p44_pedicle", to: "p44_global", direction: "forward", route: "straight", fromSide: "right", toSide: "left", color: C.ink, width: 1.25 },
  { id: "e_p44_threshold", from: "p44_global", to: "p44_threshold", direction: "forward", route: "straight", fromSide: "right", toSide: "left", color: C.ink, width: 1.25 },
];

// These are source-faithful rhetorical arrow marks with no touching endpoints;
// they are intentionally not promoted to graph edges.
const MARKERS = [
  { id: "p11_lead_arrow", geometry: "downArrow", bounds: [346, 145, 21, 47], color: C.blue },
  { id: "p11_summary_arrow", geometry: "downArrow", bounds: [346, 194, 21, 38], color: C.blue },
  { id: "p12_include_arrow", geometry: "rightArrow", bounds: [435, 223, 32, 28], color: C.blue },
  { id: "p21_got_arrow_1", geometry: "downArrow", bounds: [121, 529, 17, 35], color: C.green },
  { id: "p21_got_arrow_2", geometry: "downArrow", bounds: [121, 691, 17, 35], color: C.green },
];

// The source contains no photographs, microscopy, charts, screenshots, or
// composite evidence panels requiring replaceable raster insets.
const INSETS = [];

function box(bounds, extras = {}) {
  const [x, y, width, height] = bounds;
  return {
    left: x * SCALE,
    top: y * SCALE,
    width: width * SCALE,
    height: height * SCALE,
    ...extras,
  };
}

function lineConfig(color, width = 1, style = "solid") {
  return { style, fill: color, width: width * SCALE };
}

function addFrame(slide, frame) {
  const section = frame.role === "section";
  const stack = frame.role === "stack";
  return slide.shapes.add({
    geometry: "roundRect",
    name: frame.id,
    position: box(frame.bounds),
    fill: frame.fill ?? (section ? "none" : C.white),
    line: lineConfig(C.frame, section ? 1.45 : (stack ? 1.1 : 1.55), section ? "dashed" : "solid"),
    borderRadius: section ? 10 : (stack ? 6 : 8),
  });
}

function addLine(slide, item) {
  const [x1, y1] = item.from;
  const [x2, y2] = item.to;
  return slide.shapes.add({
    geometry: "line",
    name: item.id,
    position: {
      left: Math.min(x1, x2) * SCALE,
      top: Math.min(y1, y2) * SCALE,
      width: Math.max(Math.abs(x2 - x1), 0.01) * SCALE,
      height: Math.max(Math.abs(y2 - y1), 0.01) * SCALE,
      horizontalFlip: x2 < x1,
      verticalFlip: y2 < y1,
    },
    fill: "none",
    line: lineConfig(item.color ?? C.frame, item.width ?? 1.1, item.style ?? "solid"),
  });
}

function addNode(slide, node) {
  const style = STYLES[node.style];
  if (!style) throw new Error(`Unknown style ${node.style} for ${node.id}`);
  const config = {
    geometry: node.geometry ?? style.geometry,
    name: node.id,
    position: box(node.bounds, node.rotation ? { rotation: node.rotation } : {}),
    fill: node.fill ?? style.fill,
    line: lineConfig(node.line ?? style.line, node.lineWidth ?? style.lineWidth ?? 0),
  };
  if ((node.geometry ?? style.geometry) === "roundRect") {
    config.borderRadius = node.radius ?? style.radius ?? 6;
  }
  const shape = slide.shapes.add(config);
  if (node.text !== "") {
    shape.text = node.text;
    shape.text.style = {
      typeface: FONT,
      fontSize: node.fontSize ?? style.fontSize,
      color: node.color ?? C.ink,
      bold: node.bold ?? style.bold ?? false,
      alignment: node.align ?? style.align ?? "center",
      verticalAlignment: node.vAlign ?? "middle",
      lineSpacing: node.lineSpacing ?? style.lineSpacing ?? 0.96,
      wrap: "square",
      autoFit: "none",
      insets: {
        top: node.insetY ?? 1.5,
        right: node.insetX ?? 2,
        bottom: node.insetY ?? 1.5,
        left: node.insetX ?? 2,
      },
    };
  }
  return shape;
}

function connectorOptions(edge, head, tail) {
  const options = {
    kind: edge.route ?? "elbow",
    fromSide: edge.fromSide,
    toSide: edge.toSide,
    line: lineConfig(edge.color ?? C.frame, edge.width ?? 1.2, edge.lineStyle ?? "solid"),
    cap: "round",
    join: "round",
  };
  if (head) options.head = { type: "triangle", width: "sm", length: "sm" };
  if (tail) options.tail = { type: "triangle", width: "sm", length: "sm" };
  return options;
}

// Artifact Tool's exporter maps head to source/start and tail to target/end.
function connectOneWay(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.from),
    shapes.get(edge.to),
    connectorOptions(edge, false, true),
  );
}

function connectBothWays(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.from),
    shapes.get(edge.to),
    connectorOptions(edge, true, true),
  );
}

function connectExactPath(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.from),
    shapes.get(edge.to),
    connectorOptions(edge, false, false),
  );
}

function addBlockArrow(slide, item) {
  return slide.shapes.add({
    geometry: item.geometry,
    name: item.id,
    position: box(item.bounds),
    fill: item.color,
    line: lineConfig(item.color, 0.5),
  });
}

async function assertFreshOutput(outputPath) {
  try {
    await fs.access(outputPath);
    throw new Error(`Refusing to overwrite existing output: ${path.basename(outputPath)}`);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function writeBlob(outputPath, blob) {
  await fs.writeFile(outputPath, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  const sourceBytes = await fs.readFile(SOURCE_PATH);
  const isPng = sourceBytes.length > 8
    && sourceBytes[0] === 0x89
    && sourceBytes[1] === 0x50
    && sourceBytes[2] === 0x4e
    && sourceBytes[3] === 0x47;
  if (!isPng) throw new Error("Adjacent source.png is missing or is not a PNG file.");
  await assertFreshOutput(OUTPUT_PATH);

  const presentation = Presentation.create({ slideSize: SLIDE });
  const slide = presentation.slides.add();
  slide.background.fill = C.canvas;

  for (const frame of FRAMES) addFrame(slide, frame);
  for (const item of LINES) addLine(slide, item);

  const shapes = new Map();
  for (const node of NODES) shapes.set(node.id, addNode(slide, node));

  const connectors = [];
  for (const edge of EDGES.filter((item) => item.visual !== "block")) {
    let connector;
    if (edge.direction === "forward") connector = connectOneWay(slide, shapes, edge);
    else if (edge.direction === "both") connector = connectBothWays(slide, shapes, edge);
    else connector = connectExactPath(slide, shapes, edge);
    connector.bringToFront();
    connectors.push(connector);
  }
  for (const edge of EDGES.filter((item) => item.visual === "block")) addBlockArrow(slide, edge);
  for (const marker of MARKERS) addBlockArrow(slide, marker);

  // Keep all labels and nodes above relationship lines while leaving panel
  // shells behind. This mirrors the source's readable z-order.
  for (const shape of shapes.values()) shape.bringToFront();

  if (process.env.SCI_DIAGRAM_RENDER_PATH) {
    const preview = await presentation.export({ slide, format: "png", scale: 1 });
    await writeBlob(process.env.SCI_DIAGRAM_RENDER_PATH, preview);
  }
  if (process.env.SCI_DIAGRAM_LAYOUT_PATH) {
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(process.env.SCI_DIAGRAM_LAYOUT_PATH, await layout.text(), "utf8");
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PATH);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
