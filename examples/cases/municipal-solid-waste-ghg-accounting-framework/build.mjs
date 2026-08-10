// Native editable reconstruction of “城市固废管理与温室气体核算框架.png”.
// Authoring runtime: @oai/artifact-tool 2.8.39.
// Tested with Node.js 24.14.0 in the OpenAI Presentations runtime.
// Run from this folder with: node build.mjs

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE_PATH = path.join(HERE, "source.png");
const OUTPUT_PATH = path.join(HERE, "editable.pptx");

const SLIDE = { width: 1261, height: 1247 };
const FONT = "Times New Roman";

// Source-derived object map. Bounds use source-image pixels [x, y, width, height].
const NODES = [
  { id: "outer_border", text: "", type: "frame", bounds: [13, 8, 1235, 1226] },
  { id: "region_upper_left", text: "", type: "region", bounds: [32, 28, 835, 818] },
  { id: "region_scenarios", text: "", type: "region", bounds: [32, 857, 835, 357] },
  { id: "region_inventories", text: "", type: "region", bounds: [921, 29, 306, 827] },

  { id: "municipal_frame", text: "", type: "framework", bounds: [191, 44, 552, 466] },
  { id: "municipal_title", text: "Municipal solid waste management system framework", type: "heading", bounds: [226, 57, 492, 40] },

  { id: "ghg_panel", text: "", type: "layer", bounds: [259, 116, 360, 82] },
  { id: "ghg_label", text: "GHG emission layer", type: "label", bounds: [316, 125, 246, 27] },
  { id: "gas_ch4", text: "CH₄", type: "gas", bounds: [287, 156, 75, 31] },
  { id: "gas_co2", text: "CO₂", type: "gas", bounds: [400, 156, 76, 31] },
  { id: "gas_n2o", text: "N₂O", type: "gas", bounds: [514, 156, 78, 31] },
  { id: "layer_1", text: "Layer (1)", type: "layer-index", bounds: [642, 132, 72, 40] },

  { id: "treatment_panel", text: "", type: "layer", bounds: [259, 209, 360, 108] },
  { id: "treatment_label", text: "MSW treatment layer", type: "label", bounds: [318, 216, 242, 27] },
  { id: "material_recovery", text: "Material recovery facility", type: "treatment", bounds: [284, 246, 304, 30] },
  { id: "treatment_composting", text: "Composting", type: "treatment", bounds: [282, 280, 108, 29] },
  { id: "treatment_incineration", text: "Incineration", type: "treatment", bounds: [399, 280, 104, 29] },
  { id: "treatment_landfill", text: "Landfill", type: "treatment", bounds: [511, 280, 80, 29] },
  { id: "layer_2", text: "Layer (2)", type: "layer-index", bounds: [642, 245, 72, 40] },

  { id: "output_panel", text: "", type: "layer", bounds: [259, 328, 360, 78] },
  { id: "output_label", text: "Desired/undesired output layer", type: "label", bounds: [309, 337, 260, 27] },
  { id: "products", text: "Products", type: "output", bounds: [280, 369, 126, 30] },
  { id: "undesirable_outputs", text: "Undesirable outputs", type: "output", bounds: [414, 369, 182, 30] },
  { id: "layer_3", text: "Layer (3)", type: "layer-index", bounds: [642, 350, 72, 40] },

  { id: "resource_panel", text: "", type: "layer", bounds: [259, 418, 360, 76] },
  { id: "resource_label", text: "Resource factor input layer", type: "label", bounds: [318, 426, 244, 27] },
  { id: "energy", text: "Energy", type: "resource", bounds: [277, 456, 78, 31] },
  { id: "capital_investment", text: "Capital investment", type: "resource", bounds: [365, 456, 156, 31] },
  { id: "services", text: "Services", type: "resource", bounds: [530, 456, 72, 31] },
  { id: "layer_4", text: "Layer (4)", type: "layer-index", bounds: [642, 432, 72, 40] },

  { id: "cost_benefit_heading", text: "Cost-benefit analysis", type: "heading", bounds: [150, 540, 214, 40] },
  { id: "accounting_heading", text: "Accounting for GHG emission potentials", type: "heading", bounds: [484, 540, 366, 40] },
  { id: "costs", text: "Costs", type: "input", bounds: [48, 613, 87, 59] },
  { id: "benefits", text: "Benefits", type: "input", bounds: [48, 700, 87, 60] },
  { id: "parameters_panel", text: "", type: "analysis-panel", bounds: [190, 592, 247, 226] },
  { id: "parameters_title", text: "Main parameters", type: "label", bounds: [220, 599, 187, 28] },
  { id: "parameters_body", text: "Separate collection\nTranspotation\nFacility construction\nOperation and Maintenance\nResidential waste disposal fee\nProducts", type: "body", bounds: [198, 628, 232, 177] },
  { id: "accounting_panel", text: "", type: "analysis-panel", bounds: [483, 592, 367, 226] },
  { id: "accounting_bullets", text: "•  Sanitary landfills\n\n•  Simple landfills\n\n•  Composting\n\n•  Incineration", type: "body", bounds: [488, 607, 177, 190] },
  { id: "technology_circle", text: "Utilization of\ndifferent treatments\nand technologies", type: "analysis-node", bounds: [657, 617, 174, 183] },

  { id: "scenario_settings", text: "Scenario settings", type: "heading", bounds: [350, 868, 194, 37] },
  { id: "scenario_s0", text: "S₀: Current management system with no changes", type: "scenario", bounds: [259, 917, 395, 40] },
  { id: "structure_group", text: "", type: "optimization-group", bounds: [44, 978, 367, 217] },
  { id: "structure_title", text: "Processing structure optimization", type: "label", bounds: [71, 987, 314, 30] },
  { id: "sa1", text: "SA₁:\nReduce\nR₁ by 1%\nand\nincrease\nR₂ by 1%", type: "scenario-card", bounds: [51, 1021, 87, 167] },
  { id: "sa2", text: "SA₂:\nReduce\nR₂ by 1%\nand\nincrease\nR₃ by 1%", type: "scenario-card", bounds: [138, 1021, 87, 167] },
  { id: "sa3", text: "SA₃:\nReduce\nR₃ by 1%\nand\nincrease\nR₄ by 1%", type: "scenario-card", bounds: [225, 1021, 90, 167] },
  { id: "sa4", text: "SA₄:\nReduce\nR₄ by 1%\nand\nincrease\nR_c by 1%", type: "scenario-card", bounds: [315, 1021, 89, 167] },
  { id: "technology_group", text: "", type: "optimization-group", bounds: [419, 978, 420, 217] },
  { id: "technology_title", text: "Processing technology optimization", type: "label", bounds: [451, 987, 359, 30] },
  { id: "sb1", text: "SB₁:\nReduce\nR₁ by 1%\nand increase\nthe mechanical\ncomposting\nrate by 1%", type: "scenario-card", bounds: [420, 1021, 94, 167] },
  { id: "sb2", text: "SB₂:\nReduce\nR₂ by 1%\nand increase\nthe compost\npretreatment\nrate by 1%", type: "scenario-card", bounds: [515, 1021, 94, 167] },
  { id: "sb3", text: "SB₃:\nReduce\nR₃ by 1% and\nincrease the\nLFG power\ngeneration rate\nby 1%", type: "scenario-card", bounds: [610, 1021, 107, 167] },
  { id: "sb4", text: "SB₄:\nReduce\nR₄ by 1% and\nincrease the\nwaste-to-energy\nincineration rate\nby 1%", type: "scenario-card", bounds: [718, 1021, 118, 167] },

  { id: "inventory_heading", text: "Greenhouse gas emission\ninventories", type: "heading", bounds: [959, 51, 230, 61] },
  { id: "landfill_inventory", text: "Sanitary landfill\n\nAerobic landfill\n\nMethane combustion\n\nAnaerobic landfill after\ncompost pretreatment\n\nLFG power generation\n\nSimple landfill", type: "inventory", bounds: [954, 142, 242, 350] },
  { id: "compost_inventory", text: "Composting\n\nCompost pretreatment", type: "inventory", bounds: [954, 505, 242, 110] },
  { id: "incineration_inventory", text: "Incineration\n\nWaste-to-energy incineration", type: "inventory", bounds: [954, 627, 242, 117] },
  { id: "cost_inventory_heading", text: "Cost-benefit analysis\ninventories", type: "heading", bounds: [972, 782, 201, 56] },
  { id: "final_assessment", text: "Multi-contextual\nMSW eco-efficiency\nassessment", type: "outcome", bounds: [951, 968, 241, 173] },
];

// Directed relationships transcribed from the source. Block arrows keep their
// exact source bounds; connector edges remain attached to native endpoint shapes.
const EDGES = [
  { id: "e01_ghg_to_treatment", from: "gas_co2", to: "treatment_panel", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", style: "black" },
  { id: "e02_treatment_to_output", from: "treatment_panel", to: "output_panel", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", style: "black" },
  { id: "e03_output_to_resource", from: "output_panel", to: "resource_panel", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", style: "black" },
  { id: "e04_framework_to_cost_benefit", from: "municipal_frame", to: "cost_benefit_heading", direction: "forward", route: "elbow", fromSide: "left", toSide: "top", style: "black" },
  { id: "e05_framework_to_accounting", from: "municipal_frame", to: "accounting_heading", direction: "forward", route: "elbow", fromSide: "right", toSide: "top", style: "black" },
  { id: "e06_costs_to_parameters", from: "costs", to: "parameters_panel", direction: "forward", route: "block-right", style: "yellow-soft", bounds: [134, 626, 54, 34] },
  { id: "e07_benefits_to_parameters", from: "benefits", to: "parameters_panel", direction: "forward", route: "block-right", style: "yellow-soft", bounds: [134, 713, 54, 34] },
  { id: "e08_parameters_to_accounting", from: "parameters_panel", to: "accounting_panel", direction: "forward", route: "block-right", style: "green-soft", bounds: [436, 684, 47, 32] },
  { id: "e09_parameters_to_scenario_settings", from: "parameters_panel", to: "scenario_settings", direction: "forward", route: "elbow", fromSide: "bottom", toSide: "left", style: "black" },
  { id: "e10_accounting_to_scenario_settings", from: "accounting_panel", to: "scenario_settings", direction: "forward", route: "elbow", fromSide: "bottom", toSide: "right", style: "black" },
  { id: "e11_s0_to_structure", from: "scenario_s0", to: "structure_group", direction: "forward", route: "elbow", fromSide: "left", toSide: "top", style: "pink" },
  { id: "e12_s0_to_technology", from: "scenario_s0", to: "technology_group", direction: "forward", route: "elbow", fromSide: "right", toSide: "top", style: "pink" },
  { id: "e13_upper_left_to_inventories", from: "region_upper_left", to: "region_inventories", direction: "forward", route: "block-right", style: "yellow", bounds: [863, 417, 58, 46] },
  { id: "e14_inventory_heading_to_landfills", from: "inventory_heading", to: "landfill_inventory", direction: "forward", route: "straight", fromSide: "bottom", toSide: "top", style: "black" },
  { id: "e15_cost_inventory_to_incineration", from: "cost_inventory_heading", to: "incineration_inventory", direction: "forward", route: "straight", fromSide: "top", toSide: "bottom", style: "black" },
  { id: "e16_inventories_to_assessment", from: "region_inventories", to: "final_assessment", direction: "forward", route: "block-down", style: "yellow", bounds: [1050, 855, 46, 114] },
  { id: "e17_scenarios_to_assessment", from: "region_scenarios", to: "final_assessment", direction: "forward", route: "block-right", style: "yellow", bounds: [865, 1034, 87, 46] },
];

// No intrinsic raster insets are needed: every meaning-bearing unit is native.
const INSETS = [];

const NODE_BY_ID = new Map(NODES.map((node) => [node.id, node]));

function bounds(id) {
  const node = NODE_BY_ID.get(id);
  if (!node) throw new Error(`Unknown node: ${id}`);
  const [left, top, width, height] = node.bounds;
  return { left, top, width, height };
}

function gradient(from, to, angleDeg = 0) {
  return {
    type: "gradient",
    gradientKind: "linear",
    angleDeg,
    stops: [
      { offset: 0, color: from },
      { offset: 100000, color: to },
    ],
  };
}

const COLORS = {
  black: "#111111",
  grayLine: "#A7A7A7",
  greenLine: "#84A968",
  blueLine: "#7097C7",
  purpleLine: "#9875A6",
  yellowLine: "#E1B23B",
  inventoryLine: "#80658E",
  pinkLine: "#D991A1",
  paleGray: "#CAD3D9",
};

const presentation = Presentation.create({ slideSize: SLIDE });
const slide = presentation.slides.add();
slide.background.fill = "#FDFDFD";

const shapes = {};

function addPlainShape(id, options = {}) {
  const shape = slide.shapes.add({
    geometry: options.geometry ?? "rect",
    name: id,
    position: bounds(id),
    fill: options.fill ?? "none",
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    ...(options.borderRadius !== undefined ? { borderRadius: options.borderRadius } : {}),
    ...(options.rotation !== undefined ? { rotation: options.rotation } : {}),
  });
  shapes[id] = shape;
  return shape;
}

function addTextShape(id, options = {}) {
  const node = NODE_BY_ID.get(id);
  const shape = slide.shapes.add({
    geometry: options.geometry ?? "textbox",
    name: id,
    position: bounds(id),
    fill: options.fill ?? "none",
    line: options.line ?? { style: "solid", fill: "none", width: 0 },
    ...(options.borderRadius !== undefined ? { borderRadius: options.borderRadius } : {}),
  });
  shape.text.style = {
    fontSize: options.fontSize ?? 18,
    typeface: options.typeface ?? FONT,
    color: options.color ?? COLORS.black,
    bold: options.bold ?? false,
    alignment: options.alignment ?? "center",
    verticalAlignment: options.verticalAlignment ?? "middle",
    lineSpacing: options.lineSpacing ?? 1,
    autoFit: "none",
    wrap: "square",
    insets: options.insets ?? { left: 3, right: 3, top: 1, bottom: 1 },
  };
  shape.text = node.text;
  shapes[id] = shape;
  return shape;
}

function addOutlinedTextBox(id, options = {}) {
  return addTextShape(id, {
    geometry: options.geometry ?? "rect",
    fill: options.fill ?? "#FFFFFF",
    line: options.line ?? { style: "solid", fill: COLORS.black, width: 1.5 },
    borderRadius: options.borderRadius,
    fontSize: options.fontSize ?? 18,
    bold: options.bold ?? false,
    alignment: options.alignment ?? "center",
    verticalAlignment: options.verticalAlignment ?? "middle",
    lineSpacing: options.lineSpacing ?? 1,
    insets: options.insets,
  });
}

// Page and semantic region boundaries.
addPlainShape("outer_border", {
  geometry: "rect",
  fill: "none",
  line: { style: "solid", fill: COLORS.black, width: 1.4 },
});
for (const id of ["region_upper_left", "region_scenarios", "region_inventories"]) {
  addPlainShape(id, {
    geometry: "rect",
    fill: "none",
    line: { style: "dashed", fill: COLORS.black, width: 1.5 },
  });
}

// Large inter-module arrows and the three filled local analysis arrows are drawn
// as single native block-arrow shapes at their source-derived locations.
function addBlockArrow(edge) {
  const [left, top, width, height] = edge.bounds;
  const isDown = edge.route === "block-down";
  const palette = edge.style === "yellow-soft"
    ? { fill: gradient("#FFFDF0", "#FFF1C7"), line: "#E6C862" }
    : edge.style === "green-soft"
      ? { fill: gradient("#F0F9ED", "#D9EDD3"), line: "#B0D3A8" }
      : { fill: gradient("#FFF180", "#FFDB4D"), line: COLORS.black };
  const shape = slide.shapes.add({
    geometry: isDown ? "downArrow" : "rightArrow",
    name: edge.id,
    position: { left, top, width, height },
    fill: palette.fill,
    line: { style: "solid", fill: palette.line, width: 1.1 },
  });
  shapes[edge.id] = shape;
}

for (const edge of EDGES.filter((edge) => edge.route.startsWith("block-"))) {
  addBlockArrow(edge);
}

// Inner framework and its four layers.
addPlainShape("municipal_frame", {
  geometry: "roundRect",
  borderRadius: 6,
  fill: gradient("#FFFFFF", "#F1F6F8", 0),
  line: { style: "solid", fill: COLORS.paleGray, width: 1 },
});
addOutlinedTextBox("municipal_title", { fontSize: 19.5, bold: true, line: { style: "solid", fill: COLORS.black, width: 1.7 } });

addPlainShape("ghg_panel", {
  geometry: "roundRect",
  borderRadius: 13,
  fill: gradient("#F1F9EE", "#DDEDCF", 0),
  line: { style: "solid", fill: COLORS.greenLine, width: 1.1 },
});
addTextShape("ghg_label", { fontSize: 18 });
for (const id of ["gas_ch4", "gas_co2", "gas_n2o"]) {
  addOutlinedTextBox(id, {
    geometry: "roundRect",
    borderRadius: 4,
    fontSize: 18.5,
    line: { style: "solid", fill: "#777777", width: 0.8 },
  });
}
addTextShape("layer_1", { fontSize: 15.3, bold: true, alignment: "left", insets: { left: 0, right: 0, top: 0, bottom: 0 } });

addPlainShape("treatment_panel", {
  geometry: "roundRect",
  borderRadius: 13,
  fill: gradient("#EFF6FF", "#DCEBFB", 0),
  line: { style: "solid", fill: COLORS.blueLine, width: 1.1 },
});
addTextShape("treatment_label", { fontSize: 18 });
for (const id of ["material_recovery", "treatment_composting", "treatment_incineration", "treatment_landfill"]) {
  addOutlinedTextBox(id, {
    geometry: "roundRect",
    borderRadius: 3,
    fontSize: id === "material_recovery" ? 17.5 : 16.5,
    line: { style: "solid", fill: COLORS.blueLine, width: 0.8 },
  });
}
addTextShape("layer_2", { fontSize: 15.3, bold: true, alignment: "left", insets: { left: 0, right: 0, top: 0, bottom: 0 } });

addPlainShape("output_panel", {
  geometry: "roundRect",
  borderRadius: 12,
  fill: gradient("#F7F0F8", "#E9DBEF", 0),
  line: { style: "solid", fill: COLORS.purpleLine, width: 1.1 },
});
addTextShape("output_label", { fontSize: 17.5 });
for (const id of ["products", "undesirable_outputs"]) {
  addOutlinedTextBox(id, {
    geometry: "roundRect",
    borderRadius: 3,
    fontSize: 16.5,
    line: { style: "solid", fill: COLORS.purpleLine, width: 0.8 },
  });
}
addTextShape("layer_3", { fontSize: 15.3, bold: true, alignment: "left", insets: { left: 0, right: 0, top: 0, bottom: 0 } });

addPlainShape("resource_panel", {
  geometry: "roundRect",
  borderRadius: 12,
  fill: gradient("#FFFDF1", "#FFF1C8", 0),
  line: { style: "solid", fill: COLORS.yellowLine, width: 1.1 },
});
addTextShape("resource_label", { fontSize: 17.5 });
for (const id of ["energy", "capital_investment", "services"]) {
  addOutlinedTextBox(id, {
    geometry: "roundRect",
    borderRadius: 3,
    fontSize: id === "capital_investment" ? 16.5 : 17,
    line: { style: "solid", fill: COLORS.yellowLine, width: 0.8 },
  });
}
addTextShape("layer_4", { fontSize: 15.3, bold: true, alignment: "left", insets: { left: 0, right: 0, top: 0, bottom: 0 } });

// Cost-benefit and emissions-accounting analysis.
addOutlinedTextBox("cost_benefit_heading", { fontSize: 18, bold: true, line: { style: "solid", fill: COLORS.black, width: 1.5 } });
addOutlinedTextBox("accounting_heading", { fontSize: 18, bold: true, line: { style: "solid", fill: COLORS.black, width: 1.5 } });
for (const id of ["costs", "benefits"]) {
  addTextShape(id, {
    geometry: "roundRect",
    borderRadius: 8,
    fill: gradient("#FFFDF1", "#FFF1C8", 0),
    line: { style: "solid", fill: COLORS.yellowLine, width: 1 },
    fontSize: 18.5,
    bold: true,
  });
}
addPlainShape("parameters_panel", {
  geometry: "roundRect",
  borderRadius: 22,
  fill: gradient("#FFFDF4", "#FFF1C9", 0),
  line: { style: "solid", fill: COLORS.yellowLine, width: 1 },
});
addTextShape("parameters_title", { fontSize: 18, bold: true });
addTextShape("parameters_body", { fontSize: 17, lineSpacing: 1.12, verticalAlignment: "top", insets: { left: 2, right: 2, top: 1, bottom: 1 } });

addPlainShape("accounting_panel", {
  geometry: "rect",
  fill: gradient("#FBFDFC", "#EDF7EE", 0),
  line: { style: "solid", fill: "#D3DED4", width: 0.8 },
});
addTextShape("accounting_bullets", {
  fontSize: 17.5,
  alignment: "left",
  verticalAlignment: "top",
  lineSpacing: 1.02,
  insets: { left: 0, right: 0, top: 0, bottom: 0 },
});
addTextShape("technology_circle", {
  geometry: "ellipse",
  fill: gradient("#F1F9EF", "#D7EBCD", 0),
  line: { style: "solid", fill: "#8DBB76", width: 1 },
  fontSize: 15.5,
  lineSpacing: 0.98,
  insets: { left: 2, right: 2, top: 4, bottom: 4 },
});

// Scenario settings and optimization groups.
addOutlinedTextBox("scenario_settings", { fontSize: 17.5, bold: true, line: { style: "solid", fill: COLORS.black, width: 1.5 } });
const s0 = addTextShape("scenario_s0", {
  geometry: "roundRect",
  borderRadius: 3,
  fill: gradient("#FDF1F4", "#F7DDE4", 0),
  line: { style: "solid", fill: "#D8B2BC", width: 0.8 },
  fontSize: 17,
});
s0.text.getRange(0, 3).bold = true;

for (const id of ["structure_group", "technology_group"]) {
  addPlainShape(id, {
    geometry: "roundRect",
    borderRadius: 20,
    fill: gradient("#FDECF0", "#F8D1DA", 0),
    line: { style: "solid", fill: COLORS.pinkLine, width: 1 },
  });
}
addTextShape("structure_title", { fontSize: 17, bold: true });
addTextShape("technology_title", { fontSize: 17, bold: true });

for (const id of ["sa1", "sa2", "sa3", "sa4", "sb1", "sb2", "sb3", "sb4"]) {
  const card = addTextShape(id, {
    geometry: "roundRect",
    borderRadius: 11,
    fill: gradient("#FFFFFF", "#F4F4F4", 0),
    line: { style: "solid", fill: "#9F9F9F", width: 0.9 },
    fontSize: id.startsWith("sb") ? 13.9 : 15.3,
    lineSpacing: id.startsWith("sb") ? 0.97 : 1.02,
    insets: id.startsWith("sb")
      ? { left: 1, right: 1, top: 3, bottom: 3 }
      : { left: 3, right: 3, top: 3, bottom: 3 },
  });
  const firstLineLength = card.text.getRange(0, 4);
  firstLineLength.bold = true;
}

// Inventories and final outcome.
addOutlinedTextBox("inventory_heading", { fontSize: 19, bold: true, line: { style: "solid", fill: COLORS.black, width: 1.5 } });
for (const id of ["landfill_inventory", "compost_inventory", "incineration_inventory"]) {
  addTextShape(id, {
    geometry: "roundRect",
    borderRadius: 22,
    fill: gradient("#F3ECF6", "#E6DCEC", 0),
    line: { style: "solid", fill: COLORS.inventoryLine, width: 1.1 },
    fontSize: id === "landfill_inventory" ? 18.3 : 18,
    lineSpacing: 1,
    insets: { left: 8, right: 8, top: 8, bottom: 8 },
  });
}
addOutlinedTextBox("cost_inventory_heading", { fontSize: 17.5, bold: true, line: { style: "solid", fill: COLORS.black, width: 1.5 } });
addTextShape("final_assessment", {
  geometry: "roundRect",
  borderRadius: 22,
  fill: gradient("#FFFFFF", "#F4F4F4", 0),
  line: { style: "solid", fill: "#929292", width: 1.1 },
  fontSize: 23,
  bold: true,
  lineSpacing: 0.96,
  insets: { left: 8, right: 8, top: 8, bottom: 8 },
});

// Tested Artifact Tool direction convention: tail is the target/end arrowhead.
function connectOneWay(from, to, options = {}) {
  return slide.shapes.connect(from, to, {
    kind: options.kind ?? "elbow",
    fromSide: options.fromSide,
    toSide: options.toSide,
    line: options.line ?? { style: "solid", fill: COLORS.black, width: 1.7 },
    tail: options.tail ?? { type: "triangle", width: "sm", length: "sm" },
  });
}

function connectBothWays(from, to, options = {}) {
  return slide.shapes.connect(from, to, {
    kind: options.kind ?? "elbow",
    fromSide: options.fromSide,
    toSide: options.toSide,
    line: options.line ?? { style: "solid", fill: COLORS.black, width: 1.7 },
    head: options.head ?? { type: "triangle", width: "sm", length: "sm" },
    tail: options.tail ?? { type: "triangle", width: "sm", length: "sm" },
  });
}

function connectExactPath(from, to, options = {}) {
  return connectOneWay(from, to, options);
}

void connectBothWays; // Kept as the explicit reciprocal-edge convention; unused here.

for (const edge of EDGES.filter((edge) => !edge.route.startsWith("block-"))) {
  const line = edge.style === "pink"
    ? { style: "solid", fill: "#B97584", width: 1.4 }
    : { style: "solid", fill: COLORS.black, width: edge.route === "straight" ? 1.35 : 1.7 };
  const connector = connectExactPath(shapes[edge.from], shapes[edge.to], {
    kind: edge.route,
    fromSide: edge.fromSide,
    toSide: edge.toSide,
    line,
    tail: { type: "triangle", width: "sm", length: "sm" },
  });
  connector.name = edge.id;
  shapes[edge.id] = connector;
}

async function assertSource() {
  const bytes = await fs.readFile(SOURCE_PATH);
  if (bytes.length < 24 || bytes.toString("ascii", 1, 4) !== "PNG") {
    throw new Error("Adjacent source.png is not the expected PNG source.");
  }
  const width = bytes.readUInt32BE(16);
  const height = bytes.readUInt32BE(20);
  if (width !== SLIDE.width || height !== SLIDE.height) {
    throw new Error(`Expected source.png to be ${SLIDE.width}×${SLIDE.height}, got ${width}×${height}.`);
  }
}

async function refuseOverwrite() {
  try {
    await fs.access(OUTPUT_PATH);
  } catch (error) {
    if (error && error.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`Refusing to overwrite existing output: ${OUTPUT_PATH}`);
}

await assertSource();
await refuseOverwrite();

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(OUTPUT_PATH);

console.log(JSON.stringify({
  output: path.basename(OUTPUT_PATH),
  slideCount: presentation.slides.items.length,
  nodeCount: NODES.length,
  edgeCount: EDGES.length,
  insetCount: INSETS.length,
}, null, 2));
