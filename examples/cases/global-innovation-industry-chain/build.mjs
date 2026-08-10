/**
 * Editable reconstruction of the supplied global innovation industry-chain diagram.
 * Runtime: @oai/artifact-tool 2.8.39 (Artifact Tool backend).
 * Run from this folder with: node build.mjs
 * The adjacent source.png is read for provenance; every meaning-bearing element
 * on the slide is rebuilt as native text, shape, line, or connector objects.
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE_PATH = path.join(HERE, "source.png");
const OUTPUT_PATH = path.join(HERE, "editable.pptx");

const SOURCE = { width: 661, height: 776 };
const SCALE = 1.5;
const SLIDE = { width: SOURCE.width * SCALE, height: SOURCE.height * SCALE };
const FONT = "Songti SC";

const C = {
  white: "#FFFFFF",
  ink: "#3E4857",
  frame: "#3F3F3F",
  headerFill: "#DAE3F2",
  headerLine: "#6B879E",
  blueArrow: "#4472C4",
  blueStage: "#91ABDA",
  blueStageLine: "#5F7FAE",
  peachStage: "#FBE6D5",
  peachStageLine: "#D9B692",
  mint: "#C5F3E8",
  mintLine: "#308D84",
  greenHeader: "#C8E686",
  greenHeaderLine: "#82A844",
  yellowItem: "#FFF3CB",
  yellowItemLine: "#D6BE63",
  oliveArrow: "#6E7A48",
  periwinkle: "#B3BDFD",
  periwinkleLine: "#5D6BAD",
  peachItem: "#F4DFCF",
  peachItemLine: "#C6A589",
  paleBlue: "#CEE1FA",
  paleBlueLine: "#617FA3",
  purple: "#E6C5E7",
  purpleLine: "#9A649D",
  paleGreen: "#D9F2D0",
  paleGreenLine: "#5E9461",
  orangeArrow: "#DE763F",
  policyFill: "#F9EACD",
  policyLine: "#D5B454",
  goldArrow: "#E4B33D",
};

const STYLES = {
  sideHeader: { geometry: "roundRect", fill: C.headerFill, line: C.headerLine, fontSize: 10.5, radius: 3 },
  blueStage: { geometry: "chevron", fill: C.blueStage, line: C.blueStageLine, fontSize: 10.2, flip: true },
  peachStage: { geometry: "chevron", fill: C.peachStage, line: C.peachStageLine, fontSize: 10.2, flip: true },
  mintStage: { geometry: "chevron", fill: C.mint, line: C.mintLine, fontSize: 9.8, flip: true },
  method: { geometry: "rect", fill: C.blueStage, line: C.blueStageLine, fontSize: 9.3 },
  sectionTitle: { geometry: "roundRect", fill: C.mint, line: C.mintLine, fontSize: 9.2, radius: 3 },
  greenHeader: { geometry: "roundRect", fill: C.greenHeader, line: C.greenHeaderLine, fontSize: 10.1, radius: 3 },
  yellowItem: { geometry: "roundRect", fill: C.yellowItem, line: C.yellowItemLine, fontSize: 9.4, radius: 3 },
  periwinkleHeader: { geometry: "rect", fill: C.periwinkle, line: C.periwinkleLine, fontSize: 9.4 },
  peachItem: { geometry: "rect", fill: C.peachItem, line: C.peachItemLine, fontSize: 8.5 },
  peachVertical: { geometry: "rect", fill: C.peachItem, line: C.peachItemLine, fontSize: 8.1 },
  paleBlue: { geometry: "roundRect", fill: C.paleBlue, line: C.paleBlueLine, fontSize: 8.8, radius: 3 },
  purple: { geometry: "roundRect", fill: C.purple, line: C.purpleLine, fontSize: 8.7, radius: 3 },
  purpleVertical: { geometry: "rect", fill: C.purple, line: C.purpleLine, fontSize: 8.1 },
  peachHeader: { geometry: "rect", fill: C.peachItem, line: C.peachItemLine, fontSize: 8.4 },
  paleGreen: { geometry: "roundRect", fill: C.paleGreen, line: C.paleGreenLine, fontSize: 8.25, radius: 3 },
  policy: { geometry: "roundRect", fill: C.policyFill, line: C.policyLine, fontSize: 8.8, radius: 3 },
};

// Source-pixel object map. Bounds are [x, y, width, height] on the 661x776 reference.
// Vertical labels preserve normal Chinese reading order as upright stacked glyphs.
const NODES = [
  { id: "framework_header", text: "研究框架", type: "header", bounds: [19, 18, 100, 27], style: "sideHeader" },
  { id: "research_content_header", text: "研究内容", type: "header", bounds: [278, 18, 101, 27], style: "sideHeader" },
  { id: "methods_header", text: "研究方法", type: "header", bounds: [540, 18, 101, 27], style: "sideHeader" },

  { id: "stage_question", text: "提出问题", type: "stage", bounds: [20, 118, 97, 38], style: "blueStage" },
  { id: "stage_analysis", text: "分析问题", type: "stage", bounds: [20, 270, 97, 39], style: "peachStage" },
  { id: "stage_solution", text: "解决问题", type: "stage", bounds: [20, 394, 97, 37], style: "blueStage" },
  { id: "stage_mechanism", text: "影响机理研究", type: "stage", bounds: [20, 462, 97, 37], style: "mintStage", fontSize: 9.2 },
  { id: "stage_path", text: "产业链\n重构路径", type: "stage", bounds: [20, 593, 97, 39], style: "mintStage", fontSize: 9.0 },
  { id: "stage_policy", text: "政策建议", type: "stage", bounds: [20, 702, 98, 37], style: "blueStage" },

  { id: "method_1", text: "知识图谱\n专利计量\n技术预见\n区块聚类\n多维创新指数", type: "method", bounds: [542, 109, 96, 91], style: "method", lineSpacing: 0.92 },
  { id: "method_2", text: "复杂网络模型\n层级分析法\nABM模型\n压力测试\nAgent仿真模拟", type: "method", bounds: [542, 269, 96, 91], style: "method", lineSpacing: 0.92 },
  { id: "method_3", text: "多元Logi Probit\nTobit模型\nDSGE模型\n中心外围模型\n案例分析", type: "method", bounds: [542, 496, 96, 86], style: "method", fontSize: 8.75, lineSpacing: 0.9 },
  { id: "method_4", text: "SCP模型\nSWOT分析\nSPACE矩阵\nPMC指数模型", type: "method", bounds: [542, 688, 96, 69], style: "method", fontSize: 9.0, lineSpacing: 0.9 },

  { id: "section_trends", text: "全球创新趋势、创新力测度与国际比较研究", type: "section-title", bounds: [180, 58, 297, 21], style: "sectionTitle" },
  { id: "trend_header", text: "创新趋势", type: "group-title", bounds: [139, 85, 91, 24], style: "greenHeader" },
  { id: "trend_frontier", text: "前沿学科", type: "item", bounds: [139, 109, 91, 24], style: "yellowItem" },
  { id: "trend_patents", text: "专利分布", type: "item", bounds: [139, 132, 91, 23], style: "yellowItem" },
  { id: "trend_people", text: "人员交流", type: "item", bounds: [139, 154, 91, 22], style: "yellowItem" },
  { id: "trend_application", text: "产业应用", type: "item", bounds: [139, 175, 91, 23], style: "yellowItem" },
  { id: "measurement_header", text: "创新力测度", type: "group-title", bounds: [280, 85, 94, 24], style: "greenHeader" },
  { id: "measurement_horizontal", text: "“科学-技术-产业”\n横向价值链", renderText: "“科学-技术-产业”\n横向价值链", type: "item", bounds: [279, 112, 95, 45], style: "yellowItem", fontSize: 8.0, lineSpacing: 0.86 },
  { id: "measurement_vertical", text: "“实力-效力”\n纵向价值链", type: "item", bounds: [279, 159, 95, 39], style: "yellowItem", fontSize: 8.6, lineSpacing: 0.9 },
  { id: "comparison_header", text: "国际比较", type: "group-title", bounds: [428, 85, 92, 24], style: "greenHeader" },
  { id: "comparison_position", text: "地位指数", type: "item", bounds: [428, 113, 92, 28], style: "yellowItem" },
  { id: "comparison_participation", text: "参与度指数", type: "item", bounds: [428, 144, 92, 28], style: "yellowItem" },
  { id: "comparison_advantage", text: "竞争优势与劣势", type: "item", bounds: [428, 174, 92, 24], style: "yellowItem", fontSize: 8.7 },

  { id: "section_identification", text: "全球创新产业链的识别、解构与风险防范研究", type: "section-title", bounds: [180, 219, 297, 21], style: "sectionTitle", fontSize: 8.8 },
  { id: "identify_header", text: "产业链识别与解构", type: "group-title", bounds: [139, 248, 175, 22], style: "periwinkleHeader" },
  { id: "identify_key", text: "产业链关键环节", type: "item", bounds: [140, 275, 91, 21], style: "peachItem" },
  { id: "identify_network", text: "产业网络结构", type: "item", bounds: [140, 300, 91, 21], style: "peachItem" },
  { id: "identify_relation", text: "产业关联特征", type: "item", bounds: [140, 325, 91, 21], style: "peachItem" },
  { id: "identify_ripple", text: "产业链波及影响", type: "item", bounds: [140, 350, 91, 21], style: "peachItem" },
  { id: "chain_structure", text: "产业链结构特征", type: "vertical-item", bounds: [258, 273, 23, 102], style: "peachVertical", vertical: true, fontSize: 8.0, lineSpacing: 0.74 },
  { id: "reconstruction_basis", text: "为产业链重构提供现实依据", type: "vertical-item", bounds: [287, 273, 25, 102], style: "peachVertical", vertical: true, fontSize: 6.8, lineSpacing: 0.72 },
  { id: "risk_header", text: "产业链断链风险防范与化解", type: "group-title", bounds: [335, 248, 175, 22], style: "periwinkleHeader", fontSize: 8.8 },
  { id: "risk_upstream", text: "上游原料源头垄断", type: "item", bounds: [335, 275, 111, 20], style: "peachItem", fontSize: 8.2 },
  { id: "risk_process", text: "加工工艺不先进", type: "item", bounds: [335, 299, 111, 20], style: "peachItem" },
  { id: "risk_software", text: "设计软件对外依赖度高", type: "item", bounds: [335, 323, 111, 20], style: "peachItem", fontSize: 7.9 },
  { id: "risk_equipment", text: "加工设备自给率低", type: "item", bounds: [335, 347, 111, 20], style: "peachItem", fontSize: 8.2 },
  { id: "risk_supply", text: "上下游供需不匹配", type: "item", bounds: [335, 371, 111, 20], style: "peachItem", fontSize: 8.2 },
  { id: "risk_monitoring", text: "构建风险监测预警体系", type: "vertical-item", bounds: [455, 274, 28, 102], style: "peachVertical", vertical: true, fontSize: 6.9, lineSpacing: 0.72 },
  { id: "risk_path", text: "提出产业链重构最优路径", type: "vertical-item", bounds: [491, 274, 26, 102], style: "peachVertical", vertical: true, fontSize: 6.5, lineSpacing: 0.7 },

  { id: "section_mechanism", text: "全球创新产业链重构路径及影响机理研究", type: "section-title", bounds: [180, 412, 297, 23], style: "sectionTitle", fontSize: 8.8 },
  { id: "factor_relation", text: "产业关联特征", type: "item", bounds: [140, 450, 94, 18], style: "paleBlue" },
  { id: "factor_dependency", text: "产业链依存度", type: "item", bounds: [140, 471, 94, 18], style: "paleBlue" },
  { id: "factor_innovation", text: "技术创新", type: "item", bounds: [140, 492, 94, 18], style: "paleBlue" },
  { id: "factor_policy", text: "国家政策", type: "item", bounds: [140, 513, 94, 18], style: "paleBlue" },
  { id: "factor_competition", text: "国际竞争与制裁", type: "item", bounds: [140, 534, 94, 18], style: "paleBlue", fontSize: 8.25 },
  { id: "impact_factors", text: "影响因素", type: "vertical-item", bounds: [244, 468, 18, 64], style: "purpleVertical", vertical: true, fontSize: 8.1, lineSpacing: 0.84 },
  { id: "three_chains", text: "价值链\n+供应链\n+空间链", type: "mechanism-core", bounds: [279, 465, 70, 66], style: "purple", fontSize: 8.9, lineSpacing: 0.88 },
  { id: "impact_mechanism", text: "影响机理", type: "vertical-item", bounds: [370, 468, 18, 64], style: "purpleVertical", vertical: true, fontSize: 8.1, lineSpacing: 0.84 },
  { id: "outcome_boundary", text: "分工边界拓展", type: "item", bounds: [413, 450, 106, 18], style: "paleBlue" },
  { id: "outcome_cost", text: "研发与交易成本变化", type: "item", bounds: [413, 473, 106, 18], style: "paleBlue", fontSize: 8.1 },
  { id: "outcome_distribution", text: "价值分配转移", type: "item", bounds: [413, 496, 106, 18], style: "paleBlue" },
  { id: "outcome_supply", text: "供求关系变化", type: "item", bounds: [413, 519, 106, 18], style: "paleBlue" },
  { id: "mechanism_caption", text: "产业链重构影响机理", type: "caption", bounds: [265, 534, 120, 18], style: "purple", fontSize: 8.2 },

  { id: "path_elements_header", text: "产业链重构三要素", type: "group-title", bounds: [142, 572, 100, 26], style: "peachHeader", fontSize: 8.1 },
  { id: "path_value", text: "价值重构", type: "item", bounds: [142, 598, 98, 20], style: "paleGreen" },
  { id: "path_organization", text: "组织重构", type: "item", bounds: [142, 620, 98, 19], style: "paleGreen" },
  { id: "path_space", text: "空间重构", type: "item", bounds: [142, 642, 98, 19], style: "paleGreen" },
  { id: "semiconductor_header", text: "半导体产业链重构模式与路径", type: "group-title", bounds: [332, 572, 156, 27], style: "peachHeader", fontSize: 7.7 },
  { id: "mode_linkage", text: "全球产业链价值链联动", type: "item", bounds: [292, 598, 121, 20], style: "paleGreen", fontSize: 7.5 },
  { id: "mode_upgrade", text: "产业链升级", type: "item", bounds: [411, 598, 103, 20], style: "paleGreen" },
  { id: "mode_ecosystem", text: "全球产业链生态系统", type: "item", bounds: [292, 620, 121, 19], style: "paleGreen", fontSize: 7.9 },
  { id: "mode_expansion", text: "产业链横纵向拓展", type: "item", bounds: [411, 620, 103, 19], style: "paleGreen", fontSize: 7.8 },
  { id: "mode_substitution", text: "国产化替代", type: "item", bounds: [292, 642, 121, 19], style: "paleGreen" },
  { id: "mode_regional", text: "区域价值链合作", type: "item", bounds: [411, 642, 103, 19], style: "paleGreen", fontSize: 8.0 },

  { id: "section_policy", text: "我国创新产业政策优化与评价研究", type: "section-title", bounds: [180, 691, 297, 25], style: "sectionTitle", fontSize: 9.1 },
  { id: "policy_history", text: "产业政策发展\n历程及成效", type: "policy-step", bounds: [145, 718, 99, 36], style: "policy", fontSize: 8.3, lineSpacing: 0.9 },
  { id: "policy_evaluation", text: "产业政策\n评价体系", type: "policy-step", bounds: [281, 718, 92, 36], style: "policy", lineSpacing: 0.9 },
  { id: "policy_optimization", text: "产业政策\n优化", type: "policy-step", bounds: [418, 718, 90, 36], style: "policy", lineSpacing: 0.9 },
];

const FRAMES = [
  { id: "frame_trends", bounds: [123, 52, 408, 156], role: "outer" },
  { id: "frame_trend", bounds: [130, 82, 109, 121], role: "inner" },
  { id: "frame_measurement", bounds: [272, 82, 109, 121], role: "inner" },
  { id: "frame_comparison", bounds: [417, 82, 109, 121], role: "inner" },
  { id: "frame_identification", bounds: [123, 211, 408, 190], role: "outer" },
  { id: "frame_identify", bounds: [130, 242, 191, 148], role: "inner" },
  { id: "frame_risk", bounds: [326, 242, 197, 148], role: "inner" },
  { id: "frame_mechanism", bounds: [123, 405, 408, 272], role: "outer" },
  { id: "frame_mechanism_inner", bounds: [130, 437, 393, 118], role: "inner" },
  { id: "frame_path_row", bounds: [130, 568, 394, 103], role: "inner" },
  { id: "frame_path_elements", bounds: [134, 568, 115, 96], role: "inner" },
  { id: "frame_path_modes", bounds: [288, 568, 232, 96], role: "inner" },
  { id: "frame_policy", bounds: [123, 687, 408, 72], role: "outer" },
];

// Relationship map. Block-arrow edges carry source-pixel bounds for their single
// native arrow shape; connector edges carry explicit endpoints and direction.
const EDGES = [
  { id: "e_framework_question", from: "framework_header", to: "stage_question", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [52, 49, 33, 68], fill: C.blueArrow },
  { id: "e_question_analysis", from: "stage_question", to: "stage_analysis", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [52, 159, 33, 111], fill: C.blueArrow },
  { id: "e_analysis_solution", from: "stage_analysis", to: "stage_solution", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [52, 312, 33, 82], fill: C.blueArrow },
  { id: "e_solution_mechanism", from: "stage_solution", to: "stage_mechanism", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [53, 435, 31, 27], fill: C.blueArrow },
  { id: "e_mechanism_path", from: "stage_mechanism", to: "stage_path", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [52, 508, 33, 85], fill: C.blueArrow },
  { id: "e_path_policy", from: "stage_path", to: "stage_policy", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [52, 634, 33, 69], fill: C.blueArrow },

  { id: "e_methods_1", from: "methods_header", to: "method_1", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [573, 50, 34, 59], fill: C.blueArrow },
  { id: "e_methods_2", from: "method_1", to: "method_2", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [573, 204, 34, 65], fill: C.blueArrow },
  { id: "e_methods_3", from: "method_2", to: "method_3", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [570, 360, 40, 136], fill: C.blueArrow },
  { id: "e_methods_4", from: "method_3", to: "method_4", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [570, 582, 40, 106], fill: C.blueArrow },

  { id: "e_trend_measurement", from: "trend_header", to: "measurement_header", direction: "forward", route: "horizontal-block", lineStyle: "solid", visual: "block", geometry: "rightArrow", bounds: [238, 84, 33, 31], fill: C.oliveArrow },
  { id: "e_measurement_comparison", from: "measurement_header", to: "comparison_header", direction: "forward", route: "horizontal-block", lineStyle: "solid", visual: "block", geometry: "rightArrow", bounds: [382, 84, 34, 31], fill: C.oliveArrow },

  { id: "e_identify_key_structure", from: "identify_key", to: "chain_structure", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_identify_network_structure", from: "identify_network", to: "chain_structure", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_identify_relation_structure", from: "identify_relation", to: "chain_structure", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_identify_ripple_structure", from: "identify_ripple", to: "chain_structure", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_structure_basis", from: "chain_structure", to: "reconstruction_basis", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "right", toSide: "left" },

  { id: "e_risk_upstream_monitoring", from: "risk_upstream", to: "risk_monitoring", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_risk_process_monitoring", from: "risk_process", to: "risk_monitoring", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_risk_software_monitoring", from: "risk_software", to: "risk_monitoring", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_risk_equipment_monitoring", from: "risk_equipment", to: "risk_monitoring", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_risk_supply_monitoring", from: "risk_supply", to: "risk_monitoring", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_monitoring_path", from: "risk_monitoring", to: "risk_path", direction: "none", route: "straight", lineStyle: "solid", fromSide: "right", toSide: "left" },

  { id: "e_factor_relation", from: "factor_relation", to: "impact_factors", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_factor_dependency", from: "factor_dependency", to: "impact_factors", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_factor_innovation", from: "factor_innovation", to: "impact_factors", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_factor_policy", from: "factor_policy", to: "impact_factors", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_factor_competition", from: "factor_competition", to: "impact_factors", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_factors_three_chains", from: "impact_factors", to: "three_chains", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_outcome_boundary", from: "outcome_boundary", to: "impact_mechanism", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "left", toSide: "right" },
  { id: "e_outcome_cost", from: "outcome_cost", to: "impact_mechanism", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "left", toSide: "right" },
  { id: "e_outcome_distribution", from: "outcome_distribution", to: "impact_mechanism", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "left", toSide: "right" },
  { id: "e_outcome_supply", from: "outcome_supply", to: "impact_mechanism", direction: "forward", route: "elbow", lineStyle: "solid", fromSide: "left", toSide: "right" },
  { id: "e_mechanism_three_chains", from: "impact_mechanism", to: "three_chains", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "left", toSide: "right" },
  { id: "e_three_chains_caption", from: "three_chains", to: "mechanism_caption", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },

  { id: "e_elements_modes", from: "path_elements_header", to: "semiconductor_header", direction: "forward", route: "horizontal-block", lineStyle: "solid", visual: "block", geometry: "rightArrow", bounds: [252, 573, 30, 27], fill: C.orangeArrow },
  { id: "e_policy_history_evaluation", from: "policy_history", to: "policy_evaluation", direction: "forward", route: "horizontal-block", lineStyle: "solid", visual: "block", geometry: "rightArrow", bounds: [242, 722, 28, 25], fill: C.goldArrow },
  { id: "e_policy_evaluation_optimization", from: "policy_evaluation", to: "policy_optimization", direction: "forward", route: "horizontal-block", lineStyle: "solid", visual: "block", geometry: "rightArrow", bounds: [382, 722, 29, 25], fill: C.goldArrow },
];

// The source contains no photographs, microscopy, charts, or composite evidence
// that would need a replaceable local raster inset.
const INSETS = [];

function scaledBounds(bounds, extras = {}) {
  const [x, y, width, height] = bounds;
  return {
    left: x * SCALE,
    top: y * SCALE,
    width: width * SCALE,
    height: height * SCALE,
    ...extras,
  };
}

function scaledLine(color, width = 0.8, style = "solid") {
  return { style, fill: color, width: width * SCALE };
}

function stackedText(text) {
  return Array.from(text).join("\n");
}

function addFrame(slide, frame) {
  const shape = slide.shapes.add({
    geometry: "roundRect",
    name: frame.id,
    position: scaledBounds(frame.bounds),
    fill: "none",
    line: scaledLine(C.frame, frame.role === "outer" ? 1.35 : 1.15, "dashed"),
    borderRadius: 2.5 * SCALE,
  });
  return shape;
}

function addNode(slide, node) {
  const style = STYLES[node.style];
  if (!style) throw new Error(`Unknown style ${node.style} for ${node.id}`);
  const flip = node.horizontalFlip ?? style.flip ?? false;
  const position = scaledBounds(node.bounds, flip ? { horizontalFlip: true } : {});
  const config = {
    geometry: node.geometry ?? style.geometry,
    name: node.id,
    position,
    fill: node.fill ?? style.fill,
    line: scaledLine(node.line ?? style.line, node.lineWidth ?? 0.75),
  };
  if ((node.geometry ?? style.geometry) === "roundRect") {
    config.borderRadius = (node.radius ?? style.radius ?? 3) * SCALE;
  }
  const shape = slide.shapes.add(config);
  const displayText = node.renderText ?? (node.vertical ? stackedText(node.text) : node.text);
  shape.text = displayText;
  shape.text.style = {
    typeface: FONT,
    fontSize: (node.fontSize ?? style.fontSize) * SCALE,
    color: C.ink,
    bold: node.bold ?? false,
    alignment: "center",
    verticalAlignment: "middle",
    lineSpacing: node.lineSpacing ?? (node.vertical ? 0.78 : 0.92),
    wrap: "square",
    autoFit: "none",
    insets: {
      top: (node.insetY ?? 0.7) * SCALE,
      right: (node.insetX ?? 1.2) * SCALE,
      bottom: (node.insetY ?? 0.7) * SCALE,
      left: (node.insetX ?? 1.2) * SCALE,
    },
  };
  return shape;
}

function connectorConfig(edge, withHead, withTail) {
  const cfg = {
    kind: edge.route === "straight" ? "straight" : "elbow",
    fromSide: edge.fromSide,
    toSide: edge.toSide,
    line: scaledLine(C.frame, 0.72, edge.lineStyle ?? "solid"),
    cap: "round",
    join: "round",
  };
  if (withHead) cfg.head = { type: "triangle", width: "sm", length: "sm" };
  if (withTail) cfg.tail = { type: "triangle", width: "sm", length: "sm" };
  return cfg;
}

// Artifact Tool's current exporter maps head to the source/start and tail to
// the target/end. These helpers make that convention explicit and testable.
function connectOneWay(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.from),
    shapes.get(edge.to),
    connectorConfig(edge, false, true),
  );
}

function connectBothWays(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.from),
    shapes.get(edge.to),
    connectorConfig(edge, true, true),
  );
}

function connectExactPath(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.from),
    shapes.get(edge.to),
    connectorConfig(edge, false, false),
  );
}

function addBlockArrow(slide, edge) {
  return slide.shapes.add({
    geometry: edge.geometry,
    name: edge.id,
    position: scaledBounds(edge.bounds),
    fill: edge.fill,
    line: scaledLine(edge.fill, 0.45),
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
  slide.background.fill = C.white;

  for (const frame of FRAMES) addFrame(slide, frame);

  const shapes = new Map();
  for (const node of NODES) shapes.set(node.id, addNode(slide, node));

  for (const edge of EDGES.filter((item) => item.visual !== "block")) {
    if (edge.direction === "forward") connectOneWay(slide, shapes, edge);
    else if (edge.direction === "both") connectBothWays(slide, shapes, edge);
    else connectExactPath(slide, shapes, edge);
  }
  for (const edge of EDGES.filter((item) => item.visual === "block")) {
    addBlockArrow(slide, edge);
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PATH);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
