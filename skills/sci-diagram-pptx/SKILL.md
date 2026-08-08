---
name: sci-diagram-pptx
description: Reconstruct or repair user-provided scientific or academic schematic diagrams as high-fidelity, native editable PowerPoint (.pptx) files made primarily from PowerPoint shapes, text, and connectors. Use for 复刻、还原、临摹、修复或审查科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots regardless of authoring tool, especially visuals whose primary meaning is encoded by axes, scales, legends, or data-driven geometry, such as bar, line, scatter, violin, or heatmap plots; general business workflows; organization charts; commercial infographics; ordinary deck creation; redesign; OCR-only extraction; photo editing; or simple image placement.
---

# SciDiagram PPTX

Rebuild a scientific or academic schematic diagram as editable PowerPoint objects. Treat the source as the visual and content authority. Reproduce; do not redesign.

## Confirm that the source is an eligible diagram

Use this skill when meaning is carried primarily by boxes or nodes, arrows or connectors, labels, mathematical expressions, nested regions, symbolic illustrations, topology, or other explicit relationships. Typical sources include research frameworks, technical routes, process and mechanism diagrams, algorithm flows, system architectures, conceptual models, structured scientific infographics, and formula-bearing academic schematics.

Do not use this skill when the primary evidence is encoded by axes, quantitative marks, scales, legends, or data-driven geometry. Bar, line, scatter, bubble, violin, box, density, forest, volcano, heatmap, and similar statistical plots belong in a scientific plotting workflow, even when they appear in a paper and even when Python or R produced them.

For a mixed multi-panel figure, reconstruct only the schematic panel explicitly selected by the user. Do not rebuild adjacent data-plot panels under this workflow. Return `NEEDS_TARGET_SELECTION` when the intended panel is unclear.

## Apply the fidelity contract

Preserve in this order:

1. source identity and visible content;
2. text, mathematical meaning, and semantic topology;
3. native editability and object relationships;
4. geometry, layout, and reading order;
5. typography, color, stroke, and pixel-level appearance.

Never improve wording, hierarchy, colors, spacing, or connections unless the user explicitly changes the task from reproduction to redesign.

Use native PowerPoint text, shapes, connectors, fills, strokes, and groups whenever the source element can be represented honestly. Never claim native editability for a bitmap, imported SVG, path-converted text, OLE object, or fragmented image tiles.

## Route the request

Choose exactly one route:

- **Create**: an eligible scientific or academic diagram image exists and no editable reconstruction exists.
- **Repair**: the locked source image and an existing PPTX both exist; preserve correct native objects and repair only mismatches.
- **Audit**: the locked source image and an existing PPTX both exist; inspect the reconstruction without changing it.
- **Redesign**: stop using this skill and route to the general presentation workflow.

Treat PNG, JPEG, TIFF, WebP, or a rasterized PDF page as source images only after the eligibility check passes. For a PDF, first confirm the exact page and target panel, then rasterize it at inspection quality without altering content.

If multiple images are present and the target is not explicit, return `NEEDS_TARGET_SELECTION`. Do not guess from recency, filename, or visual prominence.

Return `NEEDS_SOURCE` when an Audit or Repair request lacks the authoritative source image or exact selected panel crop. A source-less package inspection can use the general presentation workflow, but it cannot establish reconstruction fidelity under this skill.

## Work with the presentation runtime

Use the installed `Presentations` skill alongside this skill for every PPTX create, repair, render, or inspection task. Apply these overrides for reconstruction:

- Treat the source image as explicit visual direction; do not load a default design template.
- Match source typography and density; generic minimum font-size and narrative-layout rules do not override the source.
- Use native PowerPoint shapes for the reconstruction even when a general deck workflow would prefer generated imagery.
- Do not use image generation, image search, Graphviz, or SVG as a substitute for source-faithful native reconstruction.
- Implement PPTX with `@oai/artifact-tool` from a JavaScript ES module. Do not use `python-pptx` to create or edit the deliverable.
- Use Python only for source inspection, package auditing, comparison, and QA.

Before coding, read the current Artifact Tool quick start and API documentation supplied by the `Presentations` skill. Probe undocumented capabilities before relying on them.

Load the current workspace dependencies and use the reported bundled Python executable for every helper and QA command. Do not assume the shell's active Conda, Homebrew, or system Python contains the Presentations dependencies, and do not hardcode a versioned runtime path.

Use distinct absolute paths:

```text
SCI_DIAGRAM_SKILL_DIR=<absolute path to this skill>
PRESENTATIONS_SKILL_DIR=<absolute path to the Presentations skill>
SCI_DIAGRAM_PYTHON=<workspace-dependency Python executable>
BUILD_DIR=<writable task-specific temporary directory>
SOURCE_IMAGE=<locked absolute source path>
PARENT_SOURCE_IMAGE=<optional locked parent path for a mixed figure>
FINAL_PPTX=<absolute final output path>
```

Keep manifests, plans, rendered slides, diffs, and reports under `BUILD_DIR`. Place only the final deliverable at `FINAL_PPTX`. Never overwrite an existing output silently.

## Run the reconstruction workflow

### 1. Lock and inspect the source

Resolve the exact image, schematic-panel crop, and intended canvas. For a mixed figure, first record the user's explicit panel selection and create a deterministic EXIF-oriented crop without rescaling or retouching:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/panel_crop.py" \
  "$PARENT_SOURCE_IMAGE" \
  --bbox <x> <y> <width> <height> \
  --label "<panel label or description>" \
  --selected-by "<user identity or role>" \
  --selected-at "<timezone-aware ISO-8601 timestamp>" \
  --evidence "<explicit selection evidence>" \
  --output-image "$BUILD_DIR/panel-source.png" \
  --output-manifest "$BUILD_DIR/panel-crop-manifest.json"
```

Treat the resulting crop as `SOURCE_IMAGE`. Do not invent panel-selection provenance or manually replace the generated crop. For either an independent diagram or a locked crop, record its SHA-256, dimensions, format, orientation, alpha state, and aspect ratio:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/source_preflight.py" \
  "$SOURCE_IMAGE" \
  --output "$BUILD_DIR/source-manifest.json"
```

Inspect the image at original resolution. Divide it into content regions and identify:

- nodes, containers, labels, connectors, decorations, and isolated raster content;
- text runs, mathematical expressions, line styles, colors, z-order, and groups;
- ambiguous characters, endpoints, crops, or hidden relationships.

Do not infer unreadable text, formulas, arrow directions, or off-canvas content. Return `NEEDS_CONTENT_CONFIRMATION` when ambiguity could change meaning.

### 2. Create the reconstruction contract

Read [reconstruction-contract.md](references/reconstruction-contract.md), the bundled [scene-plan schema](scripts/scene-plan.schema.json), and the compact [scene-plan example](references/scene-plan.example.json). Write `scene-plan.json` under `BUILD_DIR` with normalized source and slide coordinates, stable object and connection IDs, exact visible text, typed style and topology, expected OOXML kinds, confidence, and fallback decisions. Replace every example value with measured source evidence.

Validate it before generating any PPTX:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/validate_scene_plan.py" \
  "$BUILD_DIR/scene-plan.json" \
  --source-manifest "$BUILD_DIR/source-manifest.json" \
  --output "$BUILD_DIR/scene-plan-validation.json"
```

Do not continue while the validator reports a hard failure.

For a scene plan that declares parent/panel lineage, add `--panel-manifest "$BUILD_DIR/panel-crop-manifest.json"` to the validation command. Omitting or mismatching the deterministic crop manifest is a hard failure.

### 3. Pass the capability gate

Read [capability-matrix.md](references/capability-matrix.md) when the source contains formulas, vertical text, curved freeforms, deep groups, unusual arrows, gradients, photos, textures, or other nontrivial elements.

Classify each object:

- `native-exact`: supported without semantic or structural compromise;
- `native-approximation`: editable approximation that preserves meaning;
- `isolated-raster`: inherently raster content retained as one local image;
- `unsupported`: cannot meet the requested native-editability contract.

Require explicit user approval before using `native-approximation` for meaning-bearing content or any `isolated-raster` object. Return `NEEDS_FALLBACK_APPROVAL` or `UNSUPPORTED_NATIVE_REQUIREMENT` instead of silently degrading.

For text and formulas, read [math-and-fonts.md](references/math-and-fonts.md). Do not promise Office Math unless the generated package contains a verified native equation object and the target runtime can round-trip it safely.

### 4. Build from the scene plan

Read [native-object-policy.md](references/native-object-policy.md). Generate a task-specific `.mjs` file under `BUILD_DIR` and build from `scene-plan.json`.

Use these implementation rules:

- derive every coordinate from one source-pixel-to-slide transform;
- name objects deterministically, such as `node_001`, `edge_001`, `text_001`, and `eq_001`;
- create semantic connectors before nodes so edges remain behind labels and boxes;
- put text inside its owning shape when that preserves layout and editing quality;
- keep continuous text in coherent runs rather than one object per character;
- preserve source z-order and distinguish semantic connectors from decorative arrows;
- add isolated raster objects only when approved and never use source-image tiles to simulate editability;
- keep slide 1 free of the source image, hidden tracing images, full-slide transparent objects, and artificial page-sized background shapes;
- place the unchanged selected source, or the locked exact panel crop, on slide 2 using contain/center with no further crop;
- use identical slide size for both slides and match the locked source or panel-crop aspect ratio.

Use exactly two slides for this workflow: the native reconstruction on slide 1 and the locked source reference on slide 2. Let an explicit user filename and destination win; otherwise use a non-colliding `可编辑复现版.pptx` filename.

### 5. Render and repair

Rendering is mandatory, never optional. Render the final PPTX through a build-bound wrapper so the slide images cannot be stale or substituted:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/render_evidence.py" \
  "$FINAL_PPTX" \
  --render-script "$PRESENTATIONS_SKILL_DIR/container_tools/render_slides.py" \
  --helper-python "$SCI_DIAGRAM_PYTHON" \
  --render-dir "$BUILD_DIR/rendered" \
  --output "$BUILD_DIR/render-report.json"
```

Require a fresh render directory and a `PASS` report. Then:

1. render every slide;
2. inspect slide 1 and slide 2 individually at full size;
3. check overflow, clipping, wrapping, connector routing, z-order, and missing objects;
4. compare the slide-1 render with the source at a common canvas;
5. repair the `.mjs` source and rebuild rather than patching only the preview.

Capture a build-bound overflow report with the current `Presentations` checker:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/overflow_check.py" \
  "$FINAL_PPTX" \
  --slides-test-script "$PRESENTATIONS_SKILL_DIR/container_tools/slides_test.py" \
  --helper-python "$SCI_DIAGRAM_PYTHON" \
  --output "$BUILD_DIR/overflow-report.json"
```

After the final export, write `manual-review-attestation.json` only after actually inspecting both final slide renders at full size and confirming `overflow-report.json` reports `PASS`. Bind the attestation to the locked source, final PPTX, renders, and overflow report hashes:

```json
{
  "kind": "sci-diagram-pptx-manual-review-attestation",
  "status": "PASS",
  "source_sha256": "<locked source sha256>",
  "pptx_sha256": "<final pptx sha256>",
  "full_size_visual_review": true,
  "overflow_check_passed": true,
  "reviewed_at": "<ISO-8601 timestamp>",
  "reviewer": "<reviewing agent or person>",
  "evidence": {
    "slide_1_render": {"path": "<absolute path>", "sha256": "<sha256>"},
    "slide_2_render": {"path": "<absolute path>", "sha256": "<sha256>"},
    "overflow_report": {"path": "<absolute path to overflow-report.json>", "sha256": "<sha256>", "exit_code": 0}
  }
}
```

Never prefill either boolean or reuse an attestation from another build.

Generate comparison evidence:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/compare_render.py" \
  --reference "$SOURCE_IMAGE" \
  --render "$BUILD_DIR/rendered/slide-1.png" \
  --output-dir "$BUILD_DIR/visual-qa"
```

Use visual metrics to locate differences, not as the sole acceptance criterion. Font rasterization and antialiasing vary by renderer.

If three consecutive repair passes fail to reduce meaningful differences, return `QA_NOT_CONVERGING` with the remaining regions. Do not label the file final.

### 6. Audit native structure

Read [qa-gates.md](references/qa-gates.md). Audit the package before delivery:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/audit_pptx.py" \
  "$FINAL_PPTX" \
  --source-manifest "$BUILD_DIR/source-manifest.json" \
  --scene-plan "$BUILD_DIR/scene-plan.json" \
  --output "$BUILD_DIR/pptx-audit.json"
```

Confirm at minimum:

- the file opens and contains the expected two visible slides;
- both slides share the intended size;
- slide 1 is not a whole-slide image and contains no hidden source image;
- live text, connectors, groups, and approved raster exceptions match the plan;
- no macros, OLE packages, external media, SVG/EMF/WMF shortcuts, or orphaned slide parts exist;
- slide 2 contains the locked selected source or exact panel crop with preserved aspect ratio and no further crop;
- no unresolved placeholders, unintended overlaps, clipping, or overflow remain.

Run the aggregate gate:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/qa_gate.py" \
  --scene-plan-report "$BUILD_DIR/scene-plan-validation.json" \
  --pptx-audit "$BUILD_DIR/pptx-audit.json" \
  --render-report "$BUILD_DIR/render-report.json" \
  --visual-report "$BUILD_DIR/visual-qa/report.json" \
  --overflow-report "$BUILD_DIR/overflow-report.json" \
  --manual-review-attestation "$BUILD_DIR/manual-review-attestation.json" \
  --output "$BUILD_DIR/qa-summary.json"
```

Any unresolved hard failure blocks delivery. Deliver only when `qa-summary.json` reports `delivery_authorization: true`.

### 7. Record the verification level

Use exactly one label:

- `powerpoint-verified`: Microsoft PowerPoint opened, saved, rendered, and passed a representative editability smoke test on a temporary copy.
- `renderer-verified`: OOXML structure, standard rendering, overflow, and visual checks passed, but Microsoft PowerPoint round-trip was unavailable.

Never imply `powerpoint-verified` from LibreOffice, a thumbnail renderer, or package inspection alone.

## Handle audit and repair requests

For **Audit**, run source preflight, then build the scene plan from the locked source or import an existing trusted plan already bound to that source. Never infer the expected result from the PPTX being audited. Render, compare, audit the package, and report failures without modifying the PPTX.

For **Repair**, preserve the original file and create a copy. Reuse correct native objects, retain stable names where possible, update the scene plan, and rerun every hard gate. Do not flatten the deck to simplify repair.

## Deliver

Deliver only when the aggregate gate passes. Return the final PPTX and a concise summary containing:

- source identity;
- verification level;
- approved approximations or raster exceptions;
- any non-blocking compatibility warnings.

Do not deliver manifests, source code, renders, or QA artifacts unless requested. Do not hide unresolved uncertainty in speaker notes or invisible objects.

Use these terminal states honestly:

```text
NEEDS_TARGET_SELECTION
NEEDS_SOURCE
NEEDS_CONTENT_CONFIRMATION
NEEDS_FALLBACK_APPROVAL
UNSUPPORTED_NATIVE_REQUIREMENT
QA_NOT_CONVERGING
DELIVERED
```
