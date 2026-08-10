/**
 * Native editable reconstruction of the supplied ecosystem-services assessment flowchart.
 * Runtime: @oai/artifact-tool 2.8.39 (Artifact Tool backend).
 * Tested with Node.js 24.14.0.
 * Run from this folder with: node build.mjs
 * Use `node build.mjs --overwrite` only for an intentional rebuild.
 * The adjacent source.png is read for provenance; the slide itself contains only
 * native editable PowerPoint text, shapes, lines, and connectors.
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE_PATH = path.join(HERE, "source.png");
const OUTPUT_PATH = path.join(HERE, "editable.pptx");

const SOURCE = { width: 1695, height: 928 };
const SCALE = 720 / SOURCE.height;
const SLIDE = { width: SOURCE.width * SCALE, height: SOURCE.height * SCALE };
const FONT = "Times New Roman";
const MATH_FONT = "Cambria Math";

const C = {
  white: "#FFFFFF",
  black: "#000000",
  gray: "#C8C8C8",
  magenta: "#85003F",
  burgundy: "#820018",
  periwinkle: "#AFB7E7",
  periwinkleLine: "#8E98CE",
  blue1: "#DDF3FD",
  blue2: "#BFE5FA",
  blue3: "#CFECFB",
  lavender1: "#F3F4FB",
  lavender2: "#E0E4F3",
  peach1: "#FFF7F5",
  peach2: "#FBE9E6",
};

const FILLS = {
  ecosystemPanel: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 0,
    stops: [
      { offset: 0, color: C.blue1 },
      { offset: 50000, color: C.blue2 },
      { offset: 100000, color: C.blue1 },
    ],
  },
  blue: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 0,
    stops: [
      { offset: 0, color: C.blue3 },
      { offset: 52000, color: "#EAF8FF" },
      { offset: 100000, color: C.blue2 },
    ],
  },
  blueSoft: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 0,
    stops: [
      { offset: 0, color: "#D6EEFB" },
      { offset: 100000, color: "#C9EAFB" },
    ],
  },
  lavender: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 0,
    stops: [
      { offset: 0, color: C.lavender1 },
      { offset: 100000, color: C.lavender2 },
    ],
  },
  peach: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 0,
    stops: [
      { offset: 0, color: C.peach1 },
      { offset: 100000, color: C.peach2 },
    ],
  },
};

// Source-pixel object map. Bounds are [x, y, width, height] on the 1695x928 source.
// Literal wording and line breaks intentionally preserve the source's grammar and capitalization.
const NODES = [
  { id: "step1_prefix", text: "Step1:", type: "step-title", bounds: [36, 31, 82, 45], fontSize: 31, bold: true, italic: true, color: C.magenta, alignment: "left", insetX: 0 },
  { id: "step1_title", text: "The supply-demand balance of ESs", type: "step-title", bounds: [112, 31, 545, 45], fontSize: 31, bold: true, italic: true, alignment: "left", insetX: 0 },
  { id: "ecosystem_panel", text: "", type: "panel", bounds: [34, 97, 641, 151], fill: FILLS.ecosystemPanel, lineStyle: "dotted", lineWidth: 1.2 },
  { id: "ecosystem_panel_title", text: "Ecosystem services types", type: "text", bounds: [145, 99, 420, 34], fontSize: 28, bold: true, italic: true },
  { id: "water_yield", text: "Water\nyield", type: "service", bounds: [43, 137, 135, 75], fontSize: 25 },
  { id: "soil_conservation", text: "Soil\nConservation", type: "service", bounds: [183, 137, 157, 75], fontSize: 24 },
  { id: "carbon_sequestration", text: "Carbon\nsequestration", type: "service", bounds: [344, 137, 156, 75], fontSize: 23 },
  { id: "recreation_leisure", text: "Recreation\nand leisure", type: "service", bounds: [507, 137, 157, 75], fontSize: 23 },
  { id: "ecosystem_demand", text: "Ecosystem service\ndemand (ED)", type: "soft-flow", bounds: [64, 269, 232, 74], fontSize: 26 },
  { id: "ecosystem_supply", text: "Ecosystem service\nsupply (ES)", type: "soft-flow", bounds: [405, 269, 232, 74], fontSize: 26 },
  { id: "supply_demand_ratio", text: "Supply- demand ratio (ES/ED)", type: "ratio", bounds: [64, 400, 573, 41], fontSize: 26 },
  { id: "overall_imbalance", text: "Overall imbalance", type: "plain", bounds: [64, 489, 246, 41], fontSize: 25 },
  { id: "structural_imbalance", text: "Structural imbalance", type: "plain", bounds: [377, 489, 258, 41], fontSize: 25 },

  { id: "step2_prefix", text: "Step2:", type: "step-title", bounds: [59, 596, 88, 43], fontSize: 30, bold: true, italic: true, color: C.magenta, alignment: "left", insetX: 0 },
  { id: "step2_title", text: "The adjust factors based on  ESs scarcity", type: "step-title", bounds: [140, 596, 530, 43], fontSize: 30, bold: true, italic: true, alignment: "left", insetX: 0 },
  { id: "scarcity_box", text: "", type: "plain", bounds: [81, 661, 545, 74] },
  { id: "scarcity_title", text: "Scarcity Theory", type: "text", bounds: [158, 663, 390, 34], fontSize: 28, bold: true },
  { id: "scarcity_body", text: "The server imbalance, the higher priority urgency", type: "text", bounds: [100, 697, 507, 31], fontSize: 22 },
  { id: "uep_panel", text: "Weights among sub-regions\nUEP: the urgency index of\nsub-regional ecological\nprotection", type: "definition-panel", bounds: [25, 766, 335, 153] },
  { id: "uep_header", text: "Weights among sub-regions", type: "text", bounds: [33, 769, 319, 34], fontSize: 23 },
  { id: "uep_body", text: "UEP: the urgency index of\nsub-regional ecological\nprotection", type: "text", bounds: [38, 806, 309, 110], fontSize: 25, lineSpacing: 0.84, runs: [
    { run: "UEP:", textStyle: { bold: true } },
    { run: " the urgency index of\nsub-regional ecological\nprotection" },
  ] },
  { id: "pet_panel", text: "Weights among ESs types\nPET: the priority index\nof each ESs type in\nsub-region", type: "definition-panel", bounds: [378, 766, 307, 153] },
  { id: "pet_header", text: "Weights among ESs types", type: "text", bounds: [385, 769, 293, 34], fontSize: 23 },
  { id: "pet_body", text: "PET: the priority index\nof each ESs type in\nsub-region", type: "text", bounds: [390, 806, 281, 110], fontSize: 25, lineSpacing: 0.84, runs: [
    { run: "PET:", textStyle: { bold: true } },
    { run: " the priority index\nof each ESs type in\nsub-region" },
  ] },

  { id: "sensitivity_box", text: "Ecological Sensitivity Assessment\n(terrain, soil, climate, vegetation and human activities factors)", type: "plain", bounds: [779, 15, 574, 71] },
  { id: "sensitivity_title", text: "Ecological Sensitivity Assessment", type: "text", bounds: [806, 17, 520, 36], fontSize: 28 },
  { id: "sensitivity_body", text: "(terrain, soil, climate, vegetation and human activities factors)", type: "text", bounds: [790, 51, 552, 29], fontSize: 18 },
  { id: "step3_prefix", text: "Step3:", type: "step-title", bounds: [798, 115, 88, 39], fontSize: 30, bold: true, italic: true, color: C.magenta, alignment: "left", insetX: 0 },
  { id: "step3_title_line1", text: "Ecosystem service scarcity supply", type: "step-title", bounds: [882, 115, 453, 39], fontSize: 30, bold: true, italic: true, alignment: "left", insetX: 0 },
  { id: "step3_title_line2", text: "capacity (ESSSC)", type: "text", bounds: [880, 151, 390, 40], fontSize: 30, bold: true, italic: true },
  { id: "formula_plate", text: "ESSSC_i = UEP_j Σ^n_(m i)(PET_jm × ES_im)", type: "formula-plate", bounds: [835, 206, 480, 99] },
  { id: "formula_left", text: "ESSSCᵢ = UEPⱼ", type: "formula-text", bounds: [852, 226, 200, 54], fontSize: 31 },
  { id: "formula_sigma", text: "∑", type: "formula-text", bounds: [1048, 207, 73, 82], fontSize: 61, typeface: MATH_FONT },
  { id: "formula_upper", text: "n", type: "formula-text", bounds: [1085, 207, 31, 27], fontSize: 19 },
  { id: "formula_lower", text: "m  i", type: "formula-text", bounds: [1059, 275, 62, 23], fontSize: 18 },
  { id: "formula_right", text: "(PETⱼₘ × ESᵢₘ)", type: "formula-text", bounds: [1110, 227, 192, 54], fontSize: 30 },
  { id: "step3_explanation", text: "ESSSC in the pixel i is determined by ESs supply in the\npixel i with PET as the weight of single ESs type m in the\nsub-region j and then UEP as the weight of the total supply\nin the sub-region j.", type: "paragraph", bounds: [821, 326, 506, 91], fontSize: 21 },

  { id: "extreme_sensitive_box", text: "Extremely sensitive areas\n(the first two level of ecological\nsensitivity)", type: "criteria", bounds: [1400, 39, 284, 99] },
  { id: "extreme_sensitive_title", text: "Extremely sensitive areas", type: "text", bounds: [1411, 45, 262, 34], fontSize: 23, bold: true },
  { id: "extreme_sensitive_body", text: "(the first two level of ecological\nsensitivity)", type: "text", bounds: [1412, 78, 260, 51], fontSize: 18 },
  { id: "extreme_supply_box", text: "Extremely important\nareas of ESs supply\n(the first two level of ESSSC)", type: "criteria", bounds: [1400, 176, 284, 100] },
  { id: "extreme_supply_title", text: "Extremely important\nareas of ESs supply", type: "text", bounds: [1412, 181, 260, 59], fontSize: 23, bold: true },
  { id: "extreme_supply_body", text: "(the first two level of ESSSC)", type: "text", bounds: [1412, 241, 260, 29], fontSize: 18 },
  { id: "water_area_box", text: "Important water areas\n( > 10 km² of water area)", type: "criteria", bounds: [1400, 313, 284, 100] },
  { id: "water_area_title", text: "Important water areas", type: "text", bounds: [1412, 329, 260, 34], fontSize: 23, bold: true },
  { id: "water_area_body", text: "( > 10 km² of water area)", type: "text", bounds: [1412, 369, 260, 31], fontSize: 18 },

  { id: "step4_prefix", text: "Step4:", type: "step-title", bounds: [793, 492, 88, 45], fontSize: 31, bold: true, italic: true, color: C.magenta, alignment: "left", insetX: 0 },
  { id: "step4_title", text: "Construction of ESP", type: "step-title", bounds: [878, 492, 295, 45], fontSize: 31, bold: true, italic: true, alignment: "left", insetX: 0 },
  { id: "resistance_surface", text: "Ecological resistance\nsurface\n(resistance coefficients\nof land use types)", type: "plain", bounds: [798, 599, 289, 200] },
  { id: "resistance_title", text: "Ecological resistance\nsurface", type: "text", bounds: [814, 622, 257, 67], fontSize: 26, bold: true },
  { id: "resistance_body", text: "(resistance coefficients\nof land use types)", type: "text", bounds: [812, 705, 261, 68], fontSize: 23 },
  { id: "identify_sources", text: "Identify ecological sources", type: "action", bounds: [1197, 521, 478, 74], fontSize: 28, bold: true },
  { id: "ecological_corridors", text: "Extract ecological corridors", type: "action", bounds: [1197, 629, 478, 75], fontSize: 28, bold: true },
  { id: "ecological_nodes", text: "Extract ecological nodes", type: "action", bounds: [1197, 734, 478, 75], fontSize: 28, bold: true },
  { id: "esp_output", text: "Ecological security patterns (ESP)", type: "final", bounds: [798, 868, 877, 51], fontSize: 29, bold: true },
];

const FRAMES = [
  { id: "step1_frame", bounds: [17, 14, 675, 530], style: "dash-dot", width: 2 },
  { id: "step2_frame", bounds: [17, 583, 686, 344], style: "dash-dot", width: 2 },
  { id: "step3_frame", bounds: [779, 104, 575, 333], style: "dash-dot", width: 2 },
  { id: "criteria_frame", bounds: [1384, 14, 310, 426], style: "dotted", width: 1.2 },
  { id: "step4_frame", bounds: [776, 479, 917, 448], style: "dash-dot", width: 2 },
  { id: "step4_inner_frame", bounds: [1181, 555, 512, 275], style: "dotted", width: 1.2 },
];

// Invisible native anchor shapes preserve source-derived connector junctions without guessing routes.
const ANCHORS = [
  { id: "service_demand_join", bounds: [177, 236, 4, 4] },
  { id: "service_supply_join", bounds: [502, 236, 4, 4] },
  { id: "ratio_join", bounds: [342, 363, 4, 4] },
  { id: "overall_join", bounds: [181, 468, 4, 4] },
  { id: "structural_join", bounds: [515, 468, 4, 4] },
  { id: "scarcity_left_mid", bounds: [36, 700, 4, 4] },
  { id: "scarcity_right_mid", bounds: [668, 700, 4, 4] },
  { id: "uep_entry", bounds: [36, 764, 4, 4] },
  { id: "pet_entry", bounds: [668, 764, 4, 4] },
  { id: "step2_red_origin", bounds: [700, 775, 4, 4] },
  { id: "formula_input_upper", bounds: [833, 251, 4, 4] },
  { id: "formula_input_lower", bounds: [833, 282, 4, 4] },
  { id: "sensitivity_out", bounds: [1351, 65, 4, 4] },
  { id: "sensitivity_in", bounds: [1398, 65, 4, 4] },
  { id: "esssc_out", bounds: [1351, 224, 4, 4] },
  { id: "esssc_in", bounds: [1398, 224, 4, 4] },
  { id: "criteria_bottom", bounds: [1514, 437, 4, 4] },
  { id: "identify_top", bounds: [1514, 519, 4, 4] },
  { id: "inner_group_bottom", bounds: [1514, 827, 4, 4] },
  { id: "esp_top", bounds: [1514, 866, 4, 4] },
  { id: "mcr_start", bounds: [1085, 661, 4, 4] },
  { id: "mcr_end", bounds: [1195, 661, 4, 4] },
  { id: "mcr_lower_end", bounds: [1134, 645, 4, 4] },
  { id: "barrier_start", bounds: [1195, 691, 4, 4] },
  { id: "barrier_join", bounds: [1105, 766, 4, 4] },
  { id: "node_line_start", bounds: [1085, 766, 4, 4] },
  { id: "node_line_end", bounds: [1195, 766, 4, 4] },
];

// Relationship map. `implFrom`/`implTo` may name an invisible source-derived junction,
// while `from`/`to` retain the scientific semantic endpoints.
const EDGES = [
  { id: "e_water_trunk", from: "water_yield", to: "service_trunk", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [122, 212, 0, 26] },
  { id: "e_soil_trunk", from: "soil_conservation", to: "service_trunk", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [264, 212, 0, 26] },
  { id: "e_carbon_trunk", from: "carbon_sequestration", to: "service_trunk", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [432, 212, 0, 26] },
  { id: "e_recreation_trunk", from: "recreation_leisure", to: "service_trunk", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [555, 212, 0, 26] },
  { id: "e_service_trunk", from: "water_yield", to: "recreation_leisure", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [122, 238, 433, 0] },
  { id: "e_trunk_demand", from: "service_trunk", to: "ecosystem_demand", implFrom: "service_demand_join", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },
  { id: "e_trunk_supply", from: "service_trunk", to: "ecosystem_supply", implFrom: "service_supply_join", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },

  { id: "e_demand_merge_v", from: "ecosystem_demand", to: "ratio_join", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [178, 343, 0, 22] },
  { id: "e_supply_merge_v", from: "ecosystem_supply", to: "ratio_join", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [514, 343, 0, 22] },
  { id: "e_demand_supply_merge", from: "ecosystem_demand", to: "ecosystem_supply", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [178, 365, 336, 0] },
  { id: "e_merge_ratio", from: "ecosystem_demand+ecosystem_supply", to: "supply_demand_ratio", implFrom: "ratio_join", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },

  { id: "e_ratio_split_v", from: "supply_demand_ratio", to: "imbalance_split", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [346, 441, 0, 29] },
  { id: "e_ratio_split_h", from: "imbalance_split", to: "overall_imbalance+structural_imbalance", direction: "none", route: "exact", lineStyle: "solid", visual: "line", bounds: [183, 470, 334, 0] },
  { id: "e_split_overall", from: "supply_demand_ratio", to: "overall_imbalance", implFrom: "overall_join", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },
  { id: "e_split_structural", from: "supply_demand_ratio", to: "structural_imbalance", implFrom: "structural_join", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },

  { id: "e_overall_left_route", from: "overall_imbalance", to: "scarcity_left_mid", direction: "none", route: "elbow", lineStyle: "solid", fromSide: "left", toSide: "top" },
  { id: "e_scarcity_left", from: "scarcity_box", to: "scarcity_left_mid", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "left", toSide: "right" },
  { id: "e_left_weights", from: "overall_imbalance+scarcity_box", to: "uep_panel", implFrom: "scarcity_left_mid", implTo: "uep_entry", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },
  { id: "e_structural_right_route", from: "structural_imbalance", to: "scarcity_right_mid", direction: "none", route: "elbow", lineStyle: "solid", fromSide: "right", toSide: "top" },
  { id: "e_scarcity_right", from: "scarcity_box", to: "scarcity_right_mid", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "right", toSide: "left" },
  { id: "e_right_weights", from: "structural_imbalance+scarcity_box", to: "pet_panel", implFrom: "scarcity_right_mid", implTo: "pet_entry", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "bottom", toSide: "top" },

  { id: "e_es_formula", from: "ecosystem_supply", to: "formula_plate", implTo: "formula_input_upper", direction: "forward", route: "elbow", lineStyle: "solid", color: C.burgundy, width: 3, fromSide: "right", toSide: "left" },
  { id: "e_weights_formula", from: "step2_adjustment_factors", to: "formula_plate", implFrom: "step2_red_origin", implTo: "formula_input_lower", direction: "forward", route: "elbow", lineStyle: "solid", color: C.burgundy, width: 3, fromSide: "right", toSide: "left" },

  { id: "e_step1_step3", from: "step1_frame", to: "step3_frame", direction: "forward", route: "horizontal-block", lineStyle: "solid", visual: "block", geometry: "rightArrow", bounds: [716, 151, 37, 59], fill: C.periwinkle },
  { id: "e_step1_step2", from: "step1_frame", to: "step2_frame", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [309, 549, 51, 31], fill: C.periwinkle },
  { id: "e_step3_step4", from: "step3_frame", to: "step4_frame", direction: "forward", route: "vertical-block", lineStyle: "solid", visual: "block", geometry: "downArrow", bounds: [1047, 443, 55, 31], fill: C.periwinkle },

  { id: "e_sensitivity_criteria", from: "sensitivity_box", to: "extreme_sensitive_box", implFrom: "sensitivity_out", implTo: "sensitivity_in", direction: "forward", route: "straight", lineStyle: "solid", color: C.burgundy, width: 3, fromSide: "right", toSide: "left" },
  { id: "e_esssc_criteria", from: "formula_plate", to: "extreme_supply_box", implFrom: "esssc_out", implTo: "esssc_in", direction: "forward", route: "straight", lineStyle: "solid", color: C.burgundy, width: 3, fromSide: "right", toSide: "left" },
  { id: "e_criteria_sources", from: "criteria_frame", to: "identify_sources", implFrom: "criteria_bottom", implTo: "identify_top", direction: "forward", route: "straight", lineStyle: "solid", color: C.burgundy, width: 3, fromSide: "bottom", toSide: "top" },
  { id: "e_sources_esp", from: "step4_inner_frame", to: "esp_output", implFrom: "inner_group_bottom", implTo: "esp_top", direction: "forward", route: "straight", lineStyle: "solid", color: C.burgundy, width: 3, fromSide: "bottom", toSide: "top" },

  { id: "e_resistance_corridors", from: "resistance_surface", to: "ecological_corridors", implFrom: "mcr_start", implTo: "mcr_end", direction: "forward", route: "straight", lineStyle: "solid", label: "MCR", fromSide: "right", toSide: "left" },
  { id: "e_sources_mcr_reciprocal", from: "identify_sources", to: "mcr_junction", implTo: "mcr_lower_end", direction: "both", route: "elbow", lineStyle: "solid", fromSide: "left", toSide: "bottom" },
  { id: "e_barrier_branch", from: "ecological_corridors", to: "barrier_junction", implFrom: "barrier_start", implTo: "barrier_join", direction: "forward", route: "elbow", lineStyle: "solid", label: "Barrier point", fromSide: "left", toSide: "top" },
  { id: "e_resistance_nodes", from: "resistance_surface", to: "ecological_nodes", implFrom: "node_line_start", implTo: "node_line_end", direction: "forward", route: "straight", lineStyle: "solid", fromSide: "right", toSide: "left" },
];

// Edge labels remain labels, not semantic nodes.
const LABELS = [
  { id: "label_mcr", edge: "e_resistance_corridors", text: "MCR", bounds: [1097, 644, 92, 39], fontSize: 24, insetX: 0 },
  { id: "label_barrier", edge: "e_barrier_branch", text: "Barrier\npoint", bounds: [1108, 700, 77, 64], fontSize: 18, insetX: 0 },
];

// The source contains no photographs, microscopy, charts, screenshots, or composite evidence.
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

function scaledLine(color = C.black, width = 1.4, style = "solid") {
  return { style, fill: color, width: width * SCALE };
}

function textStyle(node, overrides = {}) {
  return {
    typeface: node.typeface ?? FONT,
    fontSize: (node.fontSize ?? 24) * SCALE,
    color: node.color ?? C.black,
    bold: node.bold ?? false,
    italic: node.italic ?? false,
    alignment: node.alignment ?? "center",
    verticalAlignment: node.verticalAlignment ?? "middle",
    lineSpacing: node.lineSpacing ?? 0.92,
    wrap: node.wrap ?? "square",
    autoFit: "none",
    insets: {
      top: (node.insetY ?? 1.5) * SCALE,
      right: (node.insetX ?? 3) * SCALE,
      bottom: (node.insetY ?? 1.5) * SCALE,
      left: (node.insetX ?? 3) * SCALE,
    },
    ...overrides,
  };
}

function nodeVisual(node) {
  switch (node.type) {
    case "panel":
      return { geometry: "rect", fill: node.fill ?? "none", line: scaledLine(C.black, node.lineWidth ?? 1.2, node.lineStyle ?? "solid") };
    case "service":
      return { geometry: "rect", fill: FILLS.blue, line: scaledLine(C.black, 1.3) };
    case "soft-flow":
      return { geometry: "rect", fill: FILLS.blueSoft, line: scaledLine(C.blue2, 0.1) };
    case "ratio":
      return { geometry: "rect", fill: FILLS.blue, line: scaledLine(C.black, 1.1) };
    case "plain":
    case "criteria":
      return { geometry: "rect", fill: C.white, line: scaledLine(C.black, 1.1) };
    case "definition-panel":
      return { geometry: "rect", fill: C.white, line: scaledLine(C.gray, 1) };
    case "formula-plate":
      return { geometry: "rect", fill: FILLS.lavender, line: scaledLine(C.black, 1.1) };
    case "action":
      return { geometry: "rect", fill: FILLS.blue, line: scaledLine(C.black, 1.2) };
    case "final":
      return { geometry: "rect", fill: FILLS.peach, line: scaledLine(C.black, 1.1) };
    default:
      return { geometry: "textbox", fill: "none", line: scaledLine("none", 0) };
  }
}

function addNode(slide, node) {
  const visual = nodeVisual(node);
  const shape = slide.shapes.add({
    geometry: visual.geometry,
    name: node.id,
    position: scaledBounds(node.bounds),
    fill: visual.fill,
    line: visual.line,
  });

  const containerOnly = ["panel", "definition-panel", "formula-plate"].includes(node.type)
    || (["plain", "criteria"].includes(node.type) && ["scarcity_box", "sensitivity_box", "resistance_surface", "extreme_sensitive_box", "extreme_supply_box", "water_area_box"].includes(node.id));
  if (containerOnly) return shape;

  if (node.type === "formula-math") {
    shape.text = [[{ latex: node.latex }]];
    shape.text.style = textStyle(node, { typeface: MATH_FONT, wrap: "none" });
    return shape;
  }

  if (node.runs) shape.text = [node.runs];
  else shape.text = node.text;
  shape.text.style = textStyle(node, node.type === "formula-text" ? { typeface: FONT, italic: true, wrap: "none" } : {});
  return shape;
}

function addFrame(slide, frame) {
  return slide.shapes.add({
    geometry: "rect",
    name: frame.id,
    position: scaledBounds(frame.bounds),
    fill: "none",
    line: scaledLine(C.black, frame.width, frame.style),
  });
}

function addAnchor(slide, anchor) {
  return slide.shapes.add({
    geometry: "rect",
    name: anchor.id,
    position: scaledBounds(anchor.bounds),
    fill: "transparent",
    line: scaledLine("none", 0),
  });
}

function connectorConfig(edge, withHead, withTail) {
  const config = {
    kind: edge.route === "straight" ? "straight" : "elbow",
    fromSide: edge.fromSide,
    toSide: edge.toSide,
    line: scaledLine(edge.color ?? C.black, edge.width ?? 1.5, edge.lineStyle ?? "solid"),
    cap: "round",
    join: "miter",
  };
  if (withHead) config.head = { type: "triangle", width: "sm", length: "sm" };
  if (withTail) config.tail = { type: "triangle", width: "sm", length: "sm" };
  return config;
}

// Artifact Tool 2.8.39 maps head to source/start and tail to target/end.
function connectOneWay(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.implFrom ?? edge.from),
    shapes.get(edge.implTo ?? edge.to),
    connectorConfig(edge, false, true),
  );
}

function connectBothWays(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.implFrom ?? edge.from),
    shapes.get(edge.implTo ?? edge.to),
    connectorConfig(edge, true, true),
  );
}

function connectExactPath(slide, shapes, edge) {
  return slide.shapes.connect(
    shapes.get(edge.implFrom ?? edge.from),
    shapes.get(edge.implTo ?? edge.to),
    connectorConfig(edge, false, false),
  );
}

function addLineSegment(slide, edge) {
  return slide.shapes.add({
    geometry: "line",
    name: edge.id,
    position: scaledBounds(edge.bounds),
    fill: "none",
    line: scaledLine(edge.color ?? C.black, edge.width ?? 1.4, edge.lineStyle ?? "solid"),
  });
}

function addBlockArrow(slide, edge) {
  return slide.shapes.add({
    geometry: edge.geometry,
    name: edge.id,
    position: scaledBounds(edge.bounds),
    fill: edge.fill,
    line: scaledLine(C.periwinkleLine, 0.8),
    shadow: "shadow-sm",
  });
}

function addDivider(slide, id, bounds) {
  return slide.shapes.add({
    geometry: "line",
    name: id,
    position: scaledBounds(bounds),
    fill: "none",
    line: scaledLine(C.gray, 1),
  });
}

function addLabel(slide, label) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: label.id,
    position: scaledBounds(label.bounds),
    fill: C.white,
    line: scaledLine("none", 0),
  });
  shape.text = label.text;
  shape.text.style = textStyle(label, { alignment: "center", verticalAlignment: "middle" });
  return shape;
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function assertSourceAndOutput() {
  const sourceBytes = await fs.readFile(SOURCE_PATH);
  const isPng = sourceBytes.length > 8
    && sourceBytes[0] === 0x89
    && sourceBytes[1] === 0x50
    && sourceBytes[2] === 0x4e
    && sourceBytes[3] === 0x47;
  if (!isPng) throw new Error("Adjacent source.png is missing or is not a PNG file.");

  if (!process.argv.includes("--overwrite")) {
    try {
      await fs.access(OUTPUT_PATH);
      throw new Error(`Refusing to overwrite existing output: ${path.basename(OUTPUT_PATH)}`);
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
    }
  }
}

async function main() {
  await assertSourceAndOutput();

  const presentation = Presentation.create({ slideSize: SLIDE });
  const slide = presentation.slides.add();
  slide.background.fill = C.white;

  for (const frame of FRAMES) addFrame(slide, frame);

  const shapes = new Map();
  for (const anchor of ANCHORS) shapes.set(anchor.id, addAnchor(slide, anchor));
  for (const node of NODES) shapes.set(node.id, addNode(slide, node));

  addDivider(slide, "uep_header_divider", [25, 804, 335, 0]);
  addDivider(slide, "pet_header_divider", [378, 804, 307, 0]);

  for (const edge of EDGES.filter((item) => item.visual === "line")) addLineSegment(slide, edge);
  for (const edge of EDGES.filter((item) => !item.visual)) {
    if (edge.direction === "forward") connectOneWay(slide, shapes, edge);
    else if (edge.direction === "both") connectBothWays(slide, shapes, edge);
    else connectExactPath(slide, shapes, edge);
  }
  for (const edge of EDGES.filter((item) => item.visual === "block")) addBlockArrow(slide, edge);
  for (const label of LABELS) addLabel(slide, label);

  const renderDir = process.env.SCI_DIAGRAM_RENDER_DIR;
  if (renderDir) {
    await fs.mkdir(renderDir, { recursive: true });
    await writeBlob(path.join(renderDir, "slide-1.png"), await presentation.export({ slide, format: "png", scale: 1 }));
    await fs.writeFile(path.join(renderDir, "slide-1.layout.json"), await (await slide.export({ format: "layout" })).text());
    const snapshot = await presentation.inspect({ kind: "slide,textbox,shape", maxChars: 30000 });
    await fs.writeFile(path.join(renderDir, "slide-1.inspect.ndjson"), snapshot.ndjson);
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PATH);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
