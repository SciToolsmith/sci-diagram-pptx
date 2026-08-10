/**
 * Scientific schematic reconstruction.
 * Runtime: @oai/artifact-tool 2.8.39 (tested with Node.js v24.14.0).
 * Run from this directory with: node build.mjs
 * The script reads adjacent source.png and refuses to overwrite editable.pptx.
 */

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SOURCE_IMAGE = path.join(HERE, "source.png");
const OUTPUT_PPTX = path.join(HERE, "editable.pptx");
const QA_DIR = process.env.SCI_DIAGRAM_QA_DIR || "";

const SOURCE = Object.freeze({ width: 1448, height: 1086 });
const SLIDE = Object.freeze({ width: 960, height: 720 });
const SCALE = SLIDE.width / SOURCE.width;
const FONT = "Times New Roman";
const BLACK = "#111111";
const BRICK = "#9B5D62";
const EXPERIMENT_RED = "#935048";
const TEAL = "#006A76";

function frame(bounds) {
  const [x1, y1, x2, y2] = bounds;
  return {
    left: x1 * SCALE,
    top: y1 * SCALE,
    width: (x2 - x1) * SCALE,
    height: (y2 - y1) * SCALE,
  };
}

function at(x, y, size = 0.5) {
  return {
    left: x * SCALE - size / 2,
    top: y * SCALE - size / 2,
    width: size,
    height: size,
  };
}

const NODES = Object.freeze([
  { id: "title", text: "water-nitrogen interaction ↔ yied, carbon emission", type: "semantic-box", bounds: [50, 19, 681, 94] },
  { id: "upper-region", text: "", type: "dashed-region", bounds: [20, 126, 708, 526] },
  { id: "nitrogen-function", text: "Building the Nitrogen\neffect function", type: "semantic-box", bounds: [95, 154, 338, 233] },
  { id: "jensen-model", text: "Jensen model", type: "semantic-box", bounds: [410, 154, 608, 233] },
  { id: "coupled-production", text: "Constructing the water-nitrogen coupled production function", type: "semantic-box", bounds: [73, 294, 666, 389] },
  { id: "solve-parameter", text: "Solving the model parameter", type: "semantic-box", bounds: [53, 415, 344, 500] },
  { id: "coupling-experiment", text: "Water-nitrogen coupling\nexperiment (maize, wheat)", type: "semantic-box-rich-text", bounds: [390, 415, 681, 500] },
  { id: "crop-model", text: "Crop suitability evaluation model", type: "container", bounds: [747, 21, 1141, 509] },
  { id: "meteorological-label", text: "Meteorological factors", type: "factor-heading", bounds: [789, 97, 1019, 122] },
  { id: "precipitation-label", text: "precipitation", type: "factor-detail", bounds: [794, 131, 923, 156] },
  { id: "temperature-label", text: "accumulated temperature", type: "factor-detail", bounds: [794, 164, 1037, 189] },
  { id: "terrain-label", text: "Terrain factors", type: "factor-heading", bounds: [789, 198, 951, 223] },
  { id: "elevation-label", text: "elevation", type: "factor-detail", bounds: [794, 235, 875, 259] },
  { id: "slope-label", text: "slope", type: "factor-detail", bounds: [794, 269, 849, 293] },
  { id: "soil-label", text: "Soil factors", type: "factor-heading", bounds: [789, 300, 916, 326] },
  { id: "soil-detail-1", text: "soil texture, pH, organic matter", type: "factor-detail", bounds: [792, 337, 1084, 362] },
  { id: "soil-detail-2", text: "total nitrogen, total phosphorus,", type: "factor-detail", bounds: [792, 371, 1100, 396] },
  { id: "soil-detail-3", text: "total potassium", type: "factor-detail", bounds: [792, 406, 950, 430] },
  { id: "management-label", text: "Management factor", type: "factor-heading", bounds: [789, 438, 1003, 465] },
  { id: "irrigation-label", text: "effective irrigation rate", type: "factor-detail", bounds: [792, 475, 1012, 501] },
  { id: "factor-analysis", text: "Factor Analysis (FA)", type: "semantic-box", bounds: [1192, 21, 1391, 99] },
  { id: "iahp", text: "Improved Analytic\nHierarchy Process\n(IAHP)", type: "semantic-box", bounds: [1192, 133, 1391, 228] },
  { id: "comprehensive-weight", text: "Comprehensive\nweight", type: "semantic-box", bounds: [1192, 278, 1392, 359] },
  { id: "crop-suitability", text: "Crop suitability (CS)", type: "semantic-box", bounds: [1192, 415, 1392, 496] },
  { id: "lower-region", text: "", type: "dashed-region", bounds: [18, 570, 703, 849] },
  { id: "state-grid-before", text: "", type: "editable-grid-5x5", bounds: [49, 598, 271, 828] },
  { id: "state-transition-top", text: "State transition", type: "edge-label", bounds: [276, 673, 442, 704] },
  { id: "state-transition-bottom", text: "function", type: "edge-label", bounds: [300, 733, 420, 765] },
  { id: "state-grid-after", text: "", type: "editable-grid-5x5", bounds: [448, 598, 670, 828] },
  { id: "optimization-panel", text: "Optimization function", type: "semantic-panel", bounds: [18, 885, 386, 1040] },
  { id: "constraints-panel", text: "Constraints", type: "semantic-panel", bounds: [418, 886, 740, 1041] },
  { id: "genetic-algorithm", text: "Genetic Algorithm\n(GA)", type: "semantic-box", bounds: [780, 609, 981, 684] },
  { id: "particle-swarm", text: "Particle swarm\nalgorithm (PSO)", type: "semantic-box", bounds: [780, 721, 981, 796] },
  { id: "optimal-results", text: "The optimal results for\ncarbon reduction", type: "semantic-box", bounds: [786, 906, 1027, 1037] },
  { id: "correlation-analysis", text: "Water and nitrogen correlation analysis", type: "semantic-box", bounds: [1073, 569, 1425, 643] },
  { id: "carbon-calculation", text: "Crop carbon emission calculation", type: "semantic-box", bounds: [1073, 674, 1425, 747] },
  { id: "water-saving-circle", text: "water\nsaving", type: "venn-circle", bounds: [1123, 783, 1276, 948] },
  { id: "nitrogen-reducing-circle", text: "nitrogen\nreducing", type: "venn-circle", bounds: [1240, 783, 1394, 948] },
  { id: "low-carbon-circle", text: "low carbon", type: "venn-circle", bounds: [1191, 872, 1341, 1038] },
  { id: "upper-merge", text: "", type: "junction", bounds: [370.5, 264.5, 371.5, 265.5] },
  { id: "method-split", text: "", type: "junction", bounds: [1173.5, 126.5, 1174.5, 127.5] },
  { id: "method-merge", text: "", type: "junction", bounds: [1409.5, 126.5, 1410.5, 127.5] },
  { id: "algorithm-split", text: "", type: "junction", bounds: [748.5, 707.5, 749.5, 708.5] },
  { id: "algorithm-merge", text: "", type: "junction", bounds: [1011.5, 840.5, 1012.5, 841.5] },
  { id: "analysis-split", text: "", type: "junction", bounds: [1043.5, 970.5, 1044.5, 971.5] },
]);

const EDGES = Object.freeze([
  { id: "e-title-upper", from: "title", to: "upper-region", direction: "forward", route: [[365, 94], [365, 127]], lineStyle: "broad-purple-arrow" },
  { id: "e-nitrogen-merge", from: "nitrogen-function", to: "upper-merge", direction: "forward", route: [[216, 233], [216, 265], [371, 265]], lineStyle: "guide-no-head" },
  { id: "e-jensen-merge", from: "jensen-model", to: "upper-merge", direction: "forward", route: [[509, 233], [509, 265], [371, 265]], lineStyle: "guide-no-head" },
  { id: "e-merge-production", from: "upper-merge", to: "coupled-production", direction: "forward", route: [[371, 265], [371, 294]], lineStyle: "one-way" },
  { id: "e-production-solve", from: "coupled-production", to: "solve-parameter", direction: "forward", route: [[199, 389], [199, 415]], lineStyle: "one-way" },
  { id: "e-experiment-solve", from: "coupling-experiment", to: "solve-parameter", direction: "forward", route: [[390, 457], [344, 457]], lineStyle: "one-way" },
  { id: "e-upper-lower", from: "upper-region", to: "lower-region", direction: "forward", route: [[354, 526], [354, 670]], lineStyle: "broad-lavender-arrow" },
  { id: "e-grid-transition", from: "state-grid-before", to: "state-grid-after", direction: "forward", route: [[271, 716], [448, 716]], lineStyle: "one-way", label: "State transition / function" },
  { id: "e-optimization-state", from: "optimization-panel", to: "lower-region", direction: "forward", route: [[205, 886], [205, 848]], lineStyle: "one-way" },
  { id: "e-constraints-state", from: "constraints-panel", to: "lower-region", direction: "forward", route: [[580, 887], [580, 849]], lineStyle: "one-way" },
  { id: "e-crop-model-methods", from: "crop-model", to: "method-split", direction: "forward", route: [[1141, 127], [1174, 127]], lineStyle: "one-way" },
  { id: "e-method-split-fa", from: "method-split", to: "factor-analysis", direction: "forward", route: [[1174, 127], [1174, 58], [1192, 58]], lineStyle: "guide-no-head" },
  { id: "e-method-split-iahp", from: "method-split", to: "iahp", direction: "forward", route: [[1174, 127], [1174, 181], [1192, 181]], lineStyle: "guide-no-head" },
  { id: "e-fa-method-merge", from: "factor-analysis", to: "method-merge", direction: "forward", route: [[1391, 59], [1410, 59], [1410, 127]], lineStyle: "guide-no-head" },
  { id: "e-iahp-method-merge", from: "iahp", to: "method-merge", direction: "forward", route: [[1391, 181], [1410, 181], [1410, 127]], lineStyle: "guide-no-head" },
  { id: "e-method-weight", from: "method-merge", to: "comprehensive-weight", direction: "forward", route: [[1410, 127], [1436, 127], [1436, 318], [1392, 318]], lineStyle: "one-way" },
  { id: "e-weight-cs", from: "comprehensive-weight", to: "crop-suitability", direction: "forward", route: [[1292, 359], [1292, 415]], lineStyle: "one-way" },
  { id: "e-cs-state", from: "crop-suitability", to: "lower-region", direction: "forward", route: [[1292, 496], [1292, 539], [732, 539], [732, 604], [703, 604]], lineStyle: "one-way" },
  { id: "e-state-algorithms", from: "lower-region", to: "algorithm-split", direction: "forward", route: [[703, 708], [749, 708]], lineStyle: "one-way" },
  { id: "e-algorithm-split-ga", from: "algorithm-split", to: "genetic-algorithm", direction: "forward", route: [[749, 708], [749, 646], [780, 646]], lineStyle: "guide-no-head" },
  { id: "e-algorithm-split-pso", from: "algorithm-split", to: "particle-swarm", direction: "forward", route: [[749, 708], [749, 758], [780, 758]], lineStyle: "guide-no-head" },
  { id: "e-ga-algorithm-merge", from: "genetic-algorithm", to: "algorithm-merge", direction: "forward", route: [[981, 646], [1012, 646], [1012, 841]], lineStyle: "guide-no-head" },
  { id: "e-pso-algorithm-merge", from: "particle-swarm", to: "algorithm-merge", direction: "forward", route: [[981, 758], [1012, 758], [1012, 841]], lineStyle: "guide-no-head" },
  { id: "e-algorithms-optimal", from: "algorithm-merge", to: "optimal-results", direction: "forward", route: [[1012, 841], [915, 841], [915, 906]], lineStyle: "one-way" },
  { id: "e-optimal-analysis-split", from: "optimal-results", to: "analysis-split", direction: "forward", route: [[1027, 971], [1044, 971], [1044, 606]], lineStyle: "guide-no-head" },
  { id: "e-analysis-correlation", from: "analysis-split", to: "correlation-analysis", direction: "forward", route: [[1044, 606], [1073, 606]], lineStyle: "one-way" },
  { id: "e-analysis-carbon", from: "analysis-split", to: "carbon-calculation", direction: "forward", route: [[1044, 606], [1044, 710], [1073, 710]], lineStyle: "one-way" },
  { id: "e-optimal-venn", from: "optimal-results", to: "venn-goal", direction: "forward", route: [[1027, 995], [1140, 995]], lineStyle: "broad-purple-arrow" },
]);

const INSETS = Object.freeze([]);

const nodeById = new Map(NODES.map((node) => [node.id, node]));
const presentation = Presentation.create({ slideSize: { width: SLIDE.width, height: SLIDE.height } });
const slide = presentation.slides.add();
slide.background.fill = "white";
slide.speakerNotes.textFrame.setText("[Sources]\n- User-provided reference image: source.png");
slide.speakerNotes.setVisible(true);

const shapes = new Map();

function addRect(id, options = {}) {
  const node = nodeById.get(id);
  if (!node) throw new Error("Unknown node: " + id);
  const shape = slide.shapes.add({
    geometry: options.geometry || "rect",
    name: id,
    position: frame(node.bounds),
    fill: options.fill === undefined ? "white" : options.fill,
    line: options.line || { style: "solid", fill: BLACK, width: 1.05 },
  });
  shapes.set(id, shape);
  if (options.text !== false && node.text) {
    shape.text = node.text;
    shape.text.style = {
      fontSize: options.fontSize || 16.2,
      typeface: FONT,
      bold: Boolean(options.bold),
      color: options.color || BLACK,
      alignment: options.alignment || "center",
      verticalAlignment: options.verticalAlignment || "middle",
      autoFit: "none",
      wrap: "square",
      insets: options.insets || { top: 2, right: 4, bottom: 2, left: 4 },
    };
  }
  return shape;
}

function addText(id, bounds, text, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: id,
    position: frame(bounds),
    fill: options.fill === undefined ? "none" : options.fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: options.fontSize || 16,
    typeface: FONT,
    bold: Boolean(options.bold),
    color: options.color || BLACK,
    alignment: options.alignment || "left",
    verticalAlignment: options.verticalAlignment || "middle",
    autoFit: "none",
    wrap: "square",
    insets: options.insets || { top: 0, right: 0, bottom: 0, left: 0 },
  };
  shapes.set(id, shape);
  return shape;
}

function addAnchor(id, x, y) {
  const anchor = slide.shapes.add({
    geometry: "ellipse",
    name: id,
    position: at(x, y),
    fill: "transparent",
    line: { style: "solid", fill: "none", width: 0 },
  });
  return anchor;
}

function connectOneWay(id, fromShape, toShape, options = {}) {
  const connector = slide.shapes.connect(fromShape, toShape, {
    kind: options.kind || "straight",
    fromSide: options.fromSide,
    toSide: options.toSide,
    line: options.line || { style: "solid", fill: BLACK, width: 1.05 },
    tail: options.tail || { type: "triangle", width: "sm", length: "sm" },
  });
  connector.name = id;
  return connector;
}

function connectBothWays(id, fromShape, toShape, options = {}) {
  const connector = slide.shapes.connect(fromShape, toShape, {
    kind: options.kind || "straight",
    fromSide: options.fromSide,
    toSide: options.toSide,
    line: options.line || { style: "solid", fill: BLACK, width: 1.05 },
    head: { type: "triangle", width: "sm", length: "sm" },
    tail: { type: "triangle", width: "sm", length: "sm" },
  });
  connector.name = id;
  return connector;
}

function connectExactPath(id, points, options = {}) {
  const anchors = points.map((point, index) => addAnchor(id + "-anchor-" + index, point[0], point[1]));
  const connectors = [];
  for (let index = 0; index < anchors.length - 1; index += 1) {
    const isLast = index === anchors.length - 2;
    const config = {
      kind: "straight",
      line: options.line || { style: "solid", fill: BLACK, width: 1.05 },
    };
    if (isLast && options.arrowAtEnd !== false) {
      config.tail = options.tail || { type: "triangle", width: "sm", length: "sm" };
    }
    if (index === 0 && options.arrowAtStart) {
      config.head = options.head || { type: "triangle", width: "sm", length: "sm" };
    }
    const connector = slide.shapes.connect(anchors[index], anchors[index + 1], config);
    connector.name = id + "-segment-" + index;
    connectors.push(connector);
  }
  return connectors;
}

function addGuidePath(id, points) {
  return connectExactPath(id, points, { arrowAtEnd: false });
}

// Structural regions and broad arrows.
addRect("upper-region", {
  text: false,
  fill: "none",
  line: { style: "dashed", fill: BLACK, width: 1.15 },
});
addRect("lower-region", {
  text: false,
  fill: "none",
  line: { style: "dashed", fill: BLACK, width: 1.15 },
});

slide.shapes.add({
  geometry: "downArrow",
  name: "e-title-upper-broad-arrow",
  position: frame([341, 94, 389, 127]),
  fill: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 90,
    stops: [
      { offset: 0, color: "#B8AFC3" },
      { offset: 100000, color: "#8F839F" },
    ],
  },
  line: { style: "solid", fill: BLACK, width: 1.0 },
});

slide.shapes.add({
  geometry: "downArrow",
  name: "e-upper-lower-broad-arrow",
  position: frame([331, 526, 378, 670]),
  fill: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 90,
    stops: [
      { offset: 0, color: "#C8C0D0" },
      { offset: 100000, color: "#AFA5BC" },
    ],
  },
  line: { style: "solid", fill: BLACK, width: 1.0 },
});

slide.shapes.add({
  geometry: "rightArrow",
  name: "e-optimal-venn-broad-arrow",
  position: frame([1027, 974, 1140, 1016]),
  fill: {
    type: "gradient",
    gradientKind: "linear",
    angleDeg: 0,
    stops: [
      { offset: 0, color: "#A79CB4" },
      { offset: 100000, color: "#8F829E" },
    ],
  },
  line: { style: "solid", fill: BLACK, width: 1.0 },
});

// Core boxes.
addRect("title", { fontSize: 18.2, bold: true, insets: { top: 3, right: 6, bottom: 3, left: 6 } });
addRect("nitrogen-function", { fontSize: 16.1 });
addRect("jensen-model", { fontSize: 16.1 });
addRect("coupled-production", { fontSize: 15.8 });
addRect("solve-parameter", { fontSize: 15.9 });

const experiment = addRect("coupling-experiment", { text: false });
experiment.text.set([
  [{ run: "Water-nitrogen coupling" }],
  [
    { run: "experiment (" },
    { run: "maize, wheat", textStyle: { color: EXPERIMENT_RED } },
    { run: ")" },
  ],
]);
experiment.text.style = {
  fontSize: 15.7,
  typeface: FONT,
  color: BLACK,
  alignment: "center",
  verticalAlignment: "middle",
  autoFit: "none",
  wrap: "square",
  insets: { top: 2, right: 3, bottom: 2, left: 3 },
};
experiment.text.get("maize, wheat").fill = EXPERIMENT_RED;

// Crop-suitability model container and its native labels.
addRect("crop-model", { text: false });
slide.shapes.add({
  geometry: "line",
  name: "crop-model-header-divider",
  position: frame([747, 86, 1141, 86]),
  fill: "none",
  line: { style: "solid", fill: BLACK, width: 1.0 },
});
addText("crop-model-title", [764, 31, 1124, 73], "Crop suitability evaluation model", {
  fontSize: 17.1,
  alignment: "center",
});

for (const [index, y] of [99, 200, 303, 441].entries()) {
  slide.shapes.add({
    geometry: "rect",
    name: "factor-marker-" + String(index + 1),
    position: frame([758, y, 775, y + 19]),
    fill: "white",
    line: { style: "solid", fill: BRICK, width: 1.35 },
  });
}

addText("meteorological-label", [789, 96, 1078, 123], "Meteorological factors", { fontSize: 16.3, color: BRICK });
addText("precipitation-label", [794, 129, 1099, 157], "precipitation", { fontSize: 16.2 });
addText("temperature-label", [794, 162, 1115, 190], "accumulated temperature", { fontSize: 16.2 });
addText("terrain-label", [789, 197, 1078, 224], "Terrain factors", { fontSize: 16.3, color: BRICK });
addText("elevation-label", [794, 232, 1100, 260], "elevation", { fontSize: 16.2 });
addText("slope-label", [794, 266, 1100, 294], "slope", { fontSize: 16.2 });
addText("soil-label", [789, 299, 1078, 327], "Soil factors", { fontSize: 16.3, color: BRICK });
addText("soil-detail-1", [792, 335, 1122, 363], "soil texture, pH, organic matter", { fontSize: 15.7 });
addText("soil-detail-2", [792, 369, 1124, 397], "total nitrogen, total phosphorus,", { fontSize: 15.7 });
addText("soil-detail-3", [792, 403, 1112, 431], "total potassium", { fontSize: 15.7 });
addText("management-label", [789, 437, 1085, 466], "Management factor", { fontSize: 16.3, color: BRICK });
addText("irrigation-label", [792, 472, 1118, 502], "effective irrigation rate", { fontSize: 15.8 });

addRect("factor-analysis", { fontSize: 13.5 });
addRect("iahp", { fontSize: 15.4 });
addRect("comprehensive-weight", { fontSize: 15.7 });
addRect("crop-suitability", { fontSize: 13.2 });

// State-transition grids: every cell is a native editable rectangle.
const leftGridColors = new Map([
  ["2-2", "#CED6B3"],
  ["2-3", "#FCD8BA"],
  ["2-4", "#FCD7B6"],
  ["3-2", "#C4CDA4"],
  ["3-3", "#B2D8E6"],
  ["3-4", "#FDE39D"],
  ["4-2", "#E9EBE0"],
  ["4-3", "#D6E9EC"],
  ["4-4", "#E9B6AF"],
]);
const rightGridColors = new Map([
  ["2-2", "#6B7F3C"],
  ["2-3", "#FBBD8C"],
  ["2-4", "#FBBB8A"],
  ["3-2", "#6E833D"],
  ["3-3", "#016C8C"],
  ["3-4", "#FDCF54"],
  ["4-2", "#D2DCBA"],
  ["4-3", "#7BBFCE"],
  ["4-4", "#D99287"],
]);

function addGrid(id, bounds, colorMap) {
  const [x1, y1, x2, y2] = bounds;
  const cellWidth = (x2 - x1) / 5;
  const cellHeight = (y2 - y1) / 5;
  for (let row = 1; row <= 5; row += 1) {
    for (let column = 1; column <= 5; column += 1) {
      const cx1 = x1 + (column - 1) * cellWidth;
      const cy1 = y1 + (row - 1) * cellHeight;
      const cx2 = x1 + column * cellWidth;
      const cy2 = y1 + row * cellHeight;
      slide.shapes.add({
        geometry: "rect",
        name: id + "-r" + row + "-c" + column,
        position: frame([cx1, cy1, cx2, cy2]),
        fill: colorMap.get(row + "-" + column) || "white",
        line: { style: "solid", fill: BLACK, width: 0.95 },
      });
    }
  }
}

addGrid("state-grid-before", nodeById.get("state-grid-before").bounds, leftGridColors);
addGrid("state-grid-after", nodeById.get("state-grid-after").bounds, rightGridColors);
addText("state-transition-top", [276, 672, 442, 704], "State transition", {
  fontSize: 15.2,
  alignment: "center",
  fill: "white",
});
addText("state-transition-bottom", [300, 734, 420, 763], "function", {
  fontSize: 15.2,
  alignment: "center",
  fill: "white",
});

// Optimization and constraints panels.
addRect("optimization-panel", { text: false });
addText("optimization-title", [42, 891, 365, 922], "Optimization function", {
  fontSize: 16.3,
  alignment: "center",
});
addText("optimization-item-1", [27, 927, 374, 955], "•  Maximum yield per unit area", {
  fontSize: 15.0,
  color: "#5E4775",
});
addText("optimization-item-2", [27, 962, 379, 990], "•  Maximum irrigation water savings", {
  fontSize: 14.6,
  color: "#8A342A",
});
addText("optimization-item-3", [27, 997, 375, 1026], "•  Minimum nitrogen input", {
  fontSize: 15.0,
  color: "#294D2A",
});

addRect("constraints-panel", { text: false });
addText("constraints-title", [438, 892, 722, 923], "Constraints", {
  fontSize: 16.3,
  alignment: "center",
});
const constraintsBody = addText("constraints-body", [430, 928, 728, 1028], "", {
  fontSize: 12.8,
  alignment: "center",
});
constraintsBody.text.set([
  [
    { run: "< ", textStyle: { color: BLACK } },
    { run: "Available agricultural water,", textStyle: { color: TEAL } },
  ],
  [{ run: "cultivated area, crop suitability,", textStyle: { color: TEAL } }],
  [
    { run: "cell irrigation water, nitrogen ", textStyle: { color: TEAL } },
    { run: ">", textStyle: { color: BLACK } },
  ],
]);
constraintsBody.text.style = {
  fontSize: 12.8,
  typeface: FONT,
  color: TEAL,
  alignment: "center",
  verticalAlignment: "middle",
  autoFit: "none",
  wrap: "square",
  insets: { top: 0, right: 0, bottom: 0, left: 0 },
};
constraintsBody.text.get("<").fill = BLACK;
constraintsBody.text.get(">").fill = BLACK;

// Algorithms, outputs, and analysis boxes.
addRect("genetic-algorithm", { fontSize: 15.7 });
addRect("particle-swarm", { fontSize: 15.5 });
addRect("optimal-results", { fontSize: 15.2 });
addRect("correlation-analysis", { fontSize: 13.0 });
addRect("carbon-calculation", { fontSize: 15.3 });

// Venn goal, with editable translucent circles and editable labels.
addRect("water-saving-circle", {
  text: false,
  geometry: "ellipse",
  fill: "#D4E5F0/68",
  line: { style: "solid", fill: BLACK, width: 1.05 },
});
addRect("nitrogen-reducing-circle", {
  text: false,
  geometry: "ellipse",
  fill: "#FDE5D5/68",
  line: { style: "solid", fill: BLACK, width: 1.05 },
});
addRect("low-carbon-circle", {
  text: false,
  geometry: "ellipse",
  fill: "#E0E9D6/68",
  line: { style: "solid", fill: BLACK, width: 1.05 },
});
addText("water-saving-label", [1152, 823, 1247, 910], "water\nsaving", {
  fontSize: 16.0,
  alignment: "center",
});
addText("nitrogen-reducing-label", [1265, 823, 1372, 910], "nitrogen\nreducing", {
  fontSize: 15.8,
  alignment: "center",
});
addText("low-carbon-label", [1210, 930, 1323, 990], "low carbon", {
  fontSize: 16.0,
  alignment: "center",
});

// Exact topology. Guide rails intentionally have no arrowhead; the shared
// group output carries the arrowhead, matching the source.
addGuidePath("e-nitrogen-merge", [[216, 233], [216, 265], [371, 265]]);
addGuidePath("e-jensen-merge", [[509, 233], [509, 265], [371, 265]]);
connectExactPath("e-merge-production", [[371, 265], [371, 294]]);
connectExactPath("e-production-solve", [[199, 389], [199, 415]]);
connectOneWay("e-experiment-solve", shapes.get("coupling-experiment"), shapes.get("solve-parameter"), {
  kind: "straight",
  fromSide: "left",
  toSide: "right",
});
connectExactPath("e-grid-transition", [[271, 716], [448, 716]]);
connectExactPath("e-optimization-state", [[205, 886], [205, 848]]);
connectExactPath("e-constraints-state", [[580, 887], [580, 849]]);

connectExactPath("e-crop-model-methods", [[1141, 127], [1174, 127]]);
addGuidePath("e-method-split-fa", [[1174, 127], [1174, 58], [1192, 58]]);
addGuidePath("e-method-split-iahp", [[1174, 127], [1174, 181], [1192, 181]]);
addGuidePath("e-fa-method-merge", [[1391, 59], [1410, 59], [1410, 127]]);
addGuidePath("e-iahp-method-merge", [[1391, 181], [1410, 181], [1410, 127]]);
connectExactPath("e-method-weight", [[1410, 127], [1436, 127], [1436, 318], [1392, 318]]);
connectOneWay("e-weight-cs", shapes.get("comprehensive-weight"), shapes.get("crop-suitability"), {
  kind: "straight",
  fromSide: "bottom",
  toSide: "top",
});
connectExactPath("e-cs-state", [[1292, 496], [1292, 539], [732, 539], [732, 604], [703, 604]]);

connectExactPath("e-state-algorithms", [[703, 708], [749, 708]]);
addGuidePath("e-algorithm-split-ga", [[749, 708], [749, 646], [780, 646]]);
addGuidePath("e-algorithm-split-pso", [[749, 708], [749, 758], [780, 758]]);
addGuidePath("e-ga-algorithm-merge", [[981, 646], [1012, 646]]);
addGuidePath("e-pso-algorithm-merge", [[981, 758], [1012, 758]]);
addGuidePath("algorithm-merge-rail", [[1012, 646], [1012, 841]]);
connectExactPath("e-algorithms-optimal", [[1012, 841], [915, 841], [915, 906]]);

addGuidePath("e-optimal-analysis-split", [[1027, 971], [1044, 971], [1044, 606]]);
connectExactPath("e-analysis-correlation", [[1044, 606], [1073, 606]]);
connectExactPath("e-analysis-carbon", [[1044, 710], [1073, 710]]);

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function ensureAbsent(filePath) {
  try {
    await fs.access(filePath);
    throw new Error("Refusing to overwrite existing output: " + filePath);
  } catch (error) {
    if (error && error.code === "ENOENT") return;
    throw error;
  }
}

async function handleInspectSidecar() {
  const sidecar = OUTPUT_PPTX + ".inspect.ndjson";
  try {
    await fs.access(sidecar);
  } catch (error) {
    if (error && error.code === "ENOENT") return;
    throw error;
  }
  if (QA_DIR) {
    await fs.mkdir(QA_DIR, { recursive: true });
    await fs.rename(sidecar, path.join(QA_DIR, "artifact-tool-export.inspect.ndjson"));
  } else {
    await fs.unlink(sidecar);
  }
}

async function main() {
  const sourceBytes = await fs.readFile(SOURCE_IMAGE);
  if (sourceBytes.length < 1024) throw new Error("Adjacent source.png is missing or invalid.");
  await ensureAbsent(OUTPUT_PPTX);

  if (QA_DIR) {
    await fs.mkdir(QA_DIR, { recursive: true });
    await writeBlob(
      path.join(QA_DIR, "artifact-tool-render.png"),
      await presentation.export({ slide, format: "png", scale: 1.5 }),
    );
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(QA_DIR, "artifact-tool-layout.json"), await layout.text());
    const inspection = await presentation.inspect({
      kind: "slide,textbox,shape,notes",
      maxChars: 50000,
    });
    await fs.writeFile(path.join(QA_DIR, "artifact-tool-inspect.ndjson"), inspection.ndjson);
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PPTX);
  await handleInspectSidecar();
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
