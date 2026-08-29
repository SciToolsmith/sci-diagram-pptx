// Portable SciDiagram PPTX build template.
// Runtime: Node.js 20+ and pptxgenjs 4.0.x.
// Replace SOURCE, NODES, EDGES, and INSETS with the current source map.
// Every visible element has exactly one owner: content retained inside an
// INSET must not also appear in NODES or EDGES.
// Run: node build.mjs [--overwrite]

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import PptxGenJS from "pptxgenjs";

const BUILD_DIR = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_PATH = path.join(BUILD_DIR, "editable.pptx");
const OVERWRITE = process.argv.includes("--overwrite");

const SOURCE = {
  file: "source.png",
  width: 1200,
  height: 675,
};

const SLIDE = {
  width: 13.333,
  height: 7.5,
};

// Standard editable objects only. Boxes are source-image pixel coordinates.
const NODES = [
  {
    id: "question",
    type: "roundRect",
    text: "Question",
    box: { x: 90, y: 235, w: 235, h: 92 },
    fill: "E8F1FF",
    line: "3B6CC0",
  },
  {
    id: "mechanism",
    type: "rect",
    text: "Mechanism",
    box: { x: 485, y: 120, w: 250, h: 92 },
    fill: "E7F7F6",
    line: "238C8C",
  },
  {
    id: "outcome",
    type: "diamond",
    text: "Outcome",
    box: { x: 485, y: 410, w: 250, h: 120 },
    fill: "F4EEFF",
    line: "7557C8",
  },
];

// The array is the topology. Points are exact source-image pixel coordinates.
// A label is an edge annotation, never a node.
const EDGES = [
  {
    id: "question-to-mechanism",
    from: "question",
    to: "mechanism",
    direction: "forward",
    points: [
      { x: 325, y: 281 },
      { x: 405, y: 281 },
      { x: 405, y: 166 },
      { x: 485, y: 166 },
    ],
    label: "frames",
    labelBox: { x: 345, y: 212, w: 120, h: 34 },
  },
  {
    id: "mechanism-outcome-feedback",
    from: "mechanism",
    to: "outcome",
    direction: "both",
    points: [
      { x: 610, y: 212 },
      { x: 610, y: 410 },
    ],
    label: "tests",
    labelBox: { x: 625, y: 290, w: 90, h: 34 },
  },
];

// Intrinsic raster evidence remains one local, replaceable image object.
const INSETS = [
  {
    id: "replaceable-image-01",
    source: "source.png",
    sourceBox: { x: 850, y: 165, w: 250, h: 180 },
    frameBox: { x: 850, y: 165, w: 250, h: 180 },
    role: "contextual",
    owns: ["local experimental trace"],
    altText: "Replaceable local source inset",
    // Set only after confirming the file has no active EXIF orientation.
    sourceOrientation: "normalized",
  },
];

const SHAPE_TYPES = {
  rect: "rect",
  roundRect: "roundRect",
  ellipse: "ellipse",
  diamond: "diamond",
};

const EDGE_DIRECTIONS = new Set(["forward", "both", "none"]);

function fail(message) {
  throw new Error(`SciDiagram portable build: ${message}`);
}

function sourcePath(relativePath) {
  if (typeof relativePath !== "string" || !relativePath || path.isAbsolute(relativePath)) {
    fail("source assets must use non-empty paths relative to build.mjs");
  }
  const resolved = path.resolve(BUILD_DIR, relativePath);
  const relative = path.relative(BUILD_DIR, resolved);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    fail(`source asset escapes the build directory: ${relativePath}`);
  }
  if (!fs.existsSync(resolved) || !fs.statSync(resolved).isFile()) {
    fail(`missing source asset: ${relativePath}`);
  }
  return resolved;
}

function toSlidePoint(point) {
  return {
    x: (point.x / SOURCE.width) * SLIDE.width,
    y: (point.y / SOURCE.height) * SLIDE.height,
  };
}

function toSlideBox(box) {
  const origin = toSlidePoint(box);
  return {
    ...origin,
    w: (box.w / SOURCE.width) * SLIDE.width,
    h: (box.h / SOURCE.height) * SLIDE.height,
  };
}

function validatePortableMap() {
  if (!(SOURCE.width > 0 && SOURCE.height > 0)) fail("SOURCE dimensions must be positive");
  if (!(SLIDE.width > 0 && SLIDE.height > 0)) fail("SLIDE dimensions must be positive");

  const nodeIds = new Set();
  for (const node of NODES) {
    if (!node.id || nodeIds.has(node.id)) fail(`duplicate or empty node id: ${node.id}`);
    if (!SHAPE_TYPES[node.type]) fail(`unsupported phase-one node type: ${node.type}`);
    nodeIds.add(node.id);
  }

  const edgeIds = new Set();
  for (const edge of EDGES) {
    if (!edge.id || edgeIds.has(edge.id)) fail(`duplicate or empty edge id: ${edge.id}`);
    if (!nodeIds.has(edge.from) || !nodeIds.has(edge.to)) {
      fail(`edge ${edge.id} references an unknown node`);
    }
    if (!EDGE_DIRECTIONS.has(edge.direction)) fail(`unsupported direction on ${edge.id}`);
    if (!Array.isArray(edge.points) || edge.points.length < 2) {
      fail(`edge ${edge.id} needs an explicit path with at least two points`);
    }
    if (edge.label && !edge.labelBox) fail(`edge ${edge.id} label needs labelBox`);
    edgeIds.add(edge.id);
  }

  const insetIds = new Set();
  for (const inset of INSETS) {
    if (!inset.id || insetIds.has(inset.id)) fail(`duplicate or empty inset id: ${inset.id}`);
    if (inset.sourceOrientation !== "normalized") {
      fail(`${inset.id} requires an orientation-normalized source image`);
    }
    for (const key of ["sourceBox", "frameBox"]) {
      const box = inset[key];
      if (!box || !(box.w > 0 && box.h > 0)) fail(`${inset.id} has an invalid ${key}`);
    }
    sourcePath(inset.source);
    insetIds.add(inset.id);
  }
}

function addDirectedSegment(slide, startPx, endPx, options = {}) {
  const start = toSlidePoint(startPx);
  const end = toSlidePoint(endPx);
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const x = Math.min(start.x, end.x);
  const y = Math.min(start.y, end.y);

  slide.addShape("line", {
    x,
    y,
    w: Math.abs(dx),
    h: Math.abs(dy),
    flipH: dx < 0,
    flipV: dy < 0,
    objectName: options.objectName,
    line: {
      color: options.color ?? "3E4652",
      width: options.width ?? 1.5,
      dashType: options.dashType ?? "solid",
      beginArrowType: options.beginArrow ? "triangle" : "none",
      endArrowType: options.endArrow ? "triangle" : "none",
    },
  });
}

function connectExactPath(slide, edge, direction = edge.direction) {
  const lastSegment = edge.points.length - 2;
  for (let index = 0; index <= lastSegment; index += 1) {
    addDirectedSegment(slide, edge.points[index], edge.points[index + 1], {
      objectName: `Edge:${edge.id}:segment-${index + 1}`,
      beginArrow: direction === "both" && index === 0,
      endArrow: (direction === "forward" || direction === "both")
        && index === lastSegment,
      color: edge.color,
      width: edge.width,
      dashType: edge.dashType,
    });
  }
}

function connectOneWay(slide, edge) {
  connectExactPath(slide, edge, "forward");
}

function connectBothWays(slide, edge) {
  connectExactPath(slide, edge, "both");
}

function addEdgeLabel(slide, edge) {
  if (!edge.label) return;
  slide.addText(edge.label, {
    ...toSlideBox(edge.labelBox),
    objectName: `Edge-label:${edge.id}`,
    fontFace: "Arial",
    fontSize: 10,
    color: "3E4652",
    margin: 0.02,
    align: "center",
    valign: "mid",
    fill: { color: "FFFFFF", transparency: 8 },
    line: { color: "FFFFFF", transparency: 100 },
    breakLine: false,
  });
}

function addInset(slide, inset) {
  const frame = toSlideBox(inset.frameBox);
  const sourceSize = inset.sourceSize ?? { width: SOURCE.width, height: SOURCE.height };
  if (inset.sourceBox.x < 0 || inset.sourceBox.y < 0
    || inset.sourceBox.x + inset.sourceBox.w > sourceSize.width
    || inset.sourceBox.y + inset.sourceBox.h > sourceSize.height) {
    fail(`${inset.id} sourceBox exceeds its source image dimensions`);
  }

  // PptxGenJS 4.0.1 expresses crop offsets against a virtual full-image box,
  // while sizing.w/h become the displayed object size. Convert pixel crop
  // coordinates into that exact contract instead of passing pixels as inches.
  const virtualFullWidth = frame.w * (sourceSize.width / inset.sourceBox.w);
  const virtualFullHeight = frame.h * (sourceSize.height / inset.sourceBox.h);
  slide.addImage({
    path: sourcePath(inset.source),
    x: frame.x,
    y: frame.y,
    w: virtualFullWidth,
    h: virtualFullHeight,
    objectName: inset.id,
    altText: inset.altText ?? `Replaceable ${inset.role ?? "local"} inset`,
    sizing: {
      type: "crop",
      x: virtualFullWidth * (inset.sourceBox.x / sourceSize.width),
      y: virtualFullHeight * (inset.sourceBox.y / sourceSize.height),
      w: frame.w,
      h: frame.h,
    },
  });
  slide.addShape("rect", {
    ...frame,
    objectName: `Inset-frame:${inset.id}`,
    fill: { color: "FFFFFF", transparency: 100 },
    line: { color: "6B7280", width: 1 },
  });
}

function addNode(slide, node) {
  const box = toSlideBox(node.box);
  slide.addShape(SHAPE_TYPES[node.type], {
    ...box,
    objectName: `Node:${node.id}`,
    fill: { color: node.fill ?? "FFFFFF" },
    line: { color: node.line ?? "5B6472", width: node.lineWidth ?? 1.25 },
  });
  slide.addText(node.text, {
    ...box,
    objectName: `Node-label:${node.id}`,
    fontFace: node.fontFace ?? "Arial",
    fontSize: node.fontSize ?? 17,
    bold: node.bold ?? true,
    color: node.color ?? "172033",
    margin: node.margin ?? 0.06,
    align: node.align ?? "center",
    valign: node.valign ?? "mid",
    breakLine: false,
  });
}

async function main() {
  if (fs.existsSync(OUTPUT_PATH) && !OVERWRITE) {
    fail("editable.pptx already exists; rerun with --overwrite only when replacement is intended");
  }
  validatePortableMap();

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SCI_SOURCE", width: SLIDE.width, height: SLIDE.height });
  pptx.layout = "SCI_SOURCE";
  pptx.author = "SciDiagram PPTX";
  pptx.subject = "Native editable scientific schematic";
  pptx.title = "SciDiagram portable reconstruction";
  pptx.lang = "en-US";
  pptx.theme = {
    headFontFace: "Arial",
    bodyFontFace: "Arial",
    lang: "en-US",
  };

  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  // Connectors first: EDGES alone define the graph and its directions.
  for (const edge of EDGES) {
    if (edge.direction === "forward") connectOneWay(slide, edge);
    else if (edge.direction === "both") connectBothWays(slide, edge);
    else connectExactPath(slide, edge, "none");
    addEdgeLabel(slide, edge);
  }
  for (const inset of INSETS) addInset(slide, inset);
  for (const node of NODES) addNode(slide, node);

  await pptx.writeFile({ fileName: OUTPUT_PATH, compression: true });
  process.stdout.write(`${JSON.stringify({ output: "editable.pptx", slides: 1, runtime: "pptxgenjs" })}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
