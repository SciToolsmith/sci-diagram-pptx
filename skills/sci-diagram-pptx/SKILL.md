---
name: sci-diagram-pptx
description: Reconstruct or repair user-provided scientific and academic schematics as visually faithful, native editable PowerPoint (.pptx) files with executable build source; also inspect existing PPTX files for meaningful fidelity and editability defects. Use for 复刻、还原、临摹、修复或检查科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots whose meaning is encoded by axes, scales, legends, or data-driven geometry; general business diagrams; organization charts; ordinary deck design; OCR-only extraction; or simple image placement.
---

# SciDiagram PPTX

Reproduce the supplied scientific schematic; do not redesign it unless the user explicitly asks.

## Preserve what matters

Prioritize, in order:

1. visible wording, symbols, formulas, and scientific meaning;
2. nodes, regions, nesting, reading order, and connector topology;
3. native editability of the meaning-bearing structure;
4. readable layout and valid export.

Treat small differences in antialiasing, kerning, exact color, line weight, corner radius, arrowhead proportions, or local spacing as cosmetic when meaning and readability remain sound. Do not chase pixel identity.

Use this skill when meaning is carried mainly by boxes, labels, formulas, nested regions, arrows, or explicit relationships. Do not use it for bar, line, scatter, violin, box, density, forest, volcano, heatmap, or similar statistical plots. For a mixed figure, reconstruct only the schematic panel selected by the user; ask which panel only when the target is unclear. Ask about unreadable content only when a character, formula, or arrow direction could change meaning.

Never silently rewrite wording, simplify topology, recolor the figure, invent a cycle from visual layout, or reverse a relationship because another design appears clearer.

## Choose one task route

- **Reconstruction**: an image or selected image panel becomes a new single-slide editable PPTX and the standard source/PPTX/build bundle.
- **Repair**: an existing PPTX is corrected without changing unrelated slides; preserve the original and return a new output with executable build source.
- **Inspection**: an existing PPTX is diagnosed without modification, build source, or a delivery folder.

Do not drift from one route into another. If a request mixes them, complete the named deliverable and report any separate inspection findings briefly.

## Select a runtime only when authoring

Choose one host-provided backend before writing `build.mjs`:

1. For Reconstruction, honor an explicit `SCI_DIAGRAM_RUNTIME=artifact-tool` or `SCI_DIAGRAM_RUNTIME=pptxgenjs` value. When unset, prefer `artifact-tool` when the installed OpenAI `Presentations` runtime is available; otherwise select `pptxgenjs` only when its package and server render dependencies are ready.
2. Repair requires a backend that can import and preserve the existing deck. Use `artifact-tool` when that capability is available. The bundled PptxGenJS phase-one route does not repair an existing deck; stop with that limitation or, with user approval, change the task to standalone slide Reconstruction.
3. Run the bundled runtime probe before authoring when availability is uncertain. Stop with a clear dependency error if the selected backend is unavailable.
4. Do not switch backend after `build.mjs` has been written, and do not place both backends in one delivered script.

Read exactly one runtime reference:

- [runtime-artifact-tool.md](references/runtime-artifact-tool.md) for Codex desktop and the OpenAI `Presentations` runtime;
- [runtime-pptxgenjs.md](references/runtime-pptxgenjs.md) for a host-configured Linux/Node.js deployment.

Inspection does not select an authoring backend or create `build.mjs`; use the available host renderer plus the bundled checker.

## Reconstruction contract

### 1. Inspect and map

Inspect the selected image at original resolution. Keep a lightweight object map inside `build.mjs` with three explicit collections when applicable:

- `NODES`: stable ID, visible text, type, and source-derived bounds;
- `EDGES`: stable ID, `from`, `to`, `direction`, route, line style, and optional label;
- `INSETS`: stable ID, source pixel box, slide frame, and semantic role.

Record relationships from the source, not from proximity or circular placement. A line label is not a node. Represent reciprocal relationships explicitly as bidirectional; never replace a pair, branch, or open path with a guessed loop.

Classify photographs, microscopy, experimental results, model screenshots, and complex composite mini-figures as intrinsic raster insets when native reconstruction would be dishonest. Keep the surrounding title, border, caption, and arrows native. Do not OCR-guess, AI-upscale, sharpen, or redraw unreadable scientific evidence. If a critical inset is too soft at its intended size, request a higher-resolution or vector asset; a contextual thumbnail may remain raster.

When the user selects a panel from a larger image, crop only that panel without resizing or retouching:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/panel_crop.py" \
  "$SOURCE_IMAGE" --bbox <x> <y> <width> <height> \
  --output "$BUILD_DIR/reference-panel.png"
```

Keep the exact upload as `SOURCE_IMAGE`. Copy it byte-for-byte into the staged folder as `source.<original-extension>` and call that copy `DELIVERY_SOURCE`. Use the original as `REFERENCE_IMAGE` unless a panel was selected; then use the temporary crop for comparison and record its displayed-pixel bounds in `build.mjs`, while still delivering the unchanged full upload.

### 2. Build native semantic units

Use the simplest coherent native object for each meaningful unit: a preset shape may own its label; a standalone label is one text box; a relationship is one connector or line; an intrinsic raster inset is one replaceable picture object. Use the selected runtime's tested `connectOneWay`, `connectBothWays`, and `connectExactPath` convention instead of raw, ambiguous arrow-end calls. Preserve origin, destination, direction, dash meaning, crossings, nesting, and z-order.

Crop an inset from `source.*` through picture crop metadata and record its source box in `INSETS`; do not deliver a separate crop file. Keep it at a readable, source-faithful size. Only add an `assets/` directory when the user separately supplies a higher-resolution or vector companion that `build.mjs` actually requires.

Treat custom geometry and freeforms as visual plates with overlaid standard text boxes unless the exported geometry contains a usable text rectangle and a focused render verifies it. Do not fragment labels into individual characters, arrows into shaft-and-head pieces, or dashed borders into many short lines. Do not use a full-slide source bitmap, hidden tracing image, imported path cloud, or image-tile mosaic on the editable slide.

Preserve the source aspect ratio and derive placement from one consistent source-to-slide transform. Create exactly one slide: the native editable reconstruction. Do not add a source-reference or hidden tracing slide.

Write `build.mjs` before export and execute that exact file to create the delivered `editable.pptx`. Import exactly one selected authoring package and only the safe built-ins accepted by the checker: `node:fs`, `node:fs/promises`, `node:path`, and `node:url`. Keep reconstruction data in the file; do not use dynamic imports, `require`, `createRequire`, or local helper modules. Resolve every companion relative to `import.meta.url`. State the selected runtime, tested version, and run command in the header. Write `editable.pptx` beside the script and refuse silent overwrite.

Read other references only when their trigger applies:

- [math-and-fonts.md](references/math-and-fonts.md): formulas, mixed scripts, true vertical text, or dense typography;
- [cross-platform-compatibility.md](references/cross-platform-compatibility.md): custom/freeform labels, fragile typography, complex formulas, a reproduced platform problem, or an explicit cross-platform request;
- [capability-matrix.md](references/capability-matrix.md): only when support for an unusual native feature is uncertain.

Use one focused capability probe when necessary. Use a close native approximation without pausing when the difference is cosmetic. Ask first when a fallback changes scientific meaning, formula content, arrow topology, or converts a substantial meaning-bearing region to raster.

### 3. Render and check

After the first export, run one render through the selected runtime and one lightweight package check. Compare the render beside the source at normal size; enlarge dense text, formulas, connector crossings, arrowheads, and raster insets only where needed. Confirm every planned edge exists once with the correct endpoint, direction, bidirectionality, label, and route. Inspect each inset at 100% for framing and readable evidence.

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$FINAL_PPTX" --source "$DELIVERY_SOURCE" --build-source "$BUILD_MJS" \
  --require-single-slide --slide 1 \
  --output "$BUILD_DIR/check.json"
```

Collect all blocking defects before editing: incorrect or guessed content; missing, extra, or reversed relationships; wrong nesting or line semantics; unreadable critical insets; clipping, overflow, serious overlap, or off-canvas content; flattened structure; corrupt export; or a package hard failure.

If blockers exist, repair them together in one focused pass, then render and run the same package check once more. After that second check, stop and report any named remaining blocker. Continue only when the user explicitly asks.

### 4. Deliver the standard folder

```text
<diagram-name>_editable/
├── source.<original-extension>
├── editable.pptx
└── build.mjs
```

The source is the unchanged upload, the PPTX contains exactly one native editable reconstruction slide, and `build.mjs` is the exact executed program. The script is reproducible in a compatible declared runtime; it is not a bundled package manager environment.

Validate the staged folder before handoff: compare source bytes, confirm the staged script was executed, and reject machine-local paths, secrets, or imports of unshipped helpers. Keep checks, renders, probes, temporary crops, caches, and intermediate exports internal. Do not add a README, manifest, preview, ZIP, `node_modules`, or package-manager files unless requested. Include `assets/` only for separately supplied companions required by the build. For several independent images, create one folder per target.

## Repair contract

Identify the exact slide number or numbers and defect before authoring; ask only when the target is unclear. Keep the original PPTX untouched and preserve unrelated slides, masters, layouts, notes, and order. Prefer a self-contained rebuild. If `build.mjs` must read the original deck, include a byte-for-byte `input-original.pptx` beside it and resolve that dependency relative to `import.meta.url`. Include a supplied reference image only when the repair or its build actually depends on it.

Return a repair folder containing `editable.pptx`, the exact executed `build.mjs`, and only the adjacent input dependencies required to rerun it. Do not describe this as the standard three-file Reconstruction bundle.

Render the repaired result and deep-check each repaired slide separately. Do not pass `--source` or `--require-single-slide` merely because the repair used a visual reference or the original deck happens to contain one slide:

```bash
TARGET_SLIDE=3  # replace with the repaired slide's 1-based number
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$REPAIRED_PPTX" --build-source "$BUILD_MJS" --slide "$TARGET_SLIDE" \
  --output "$BUILD_DIR/check-slide-$TARGET_SLIDE.json"
```

For several repaired slides, run the command once per repaired slide. Collect blockers across those results, make at most one focused correction pass, rerun only the affected slide checks, then stop and report any remaining blocker unless the user explicitly asks to continue.

## Inspection contract

Do not modify the file, select an authoring backend, create `build.mjs`, or create a delivery folder. If the deck has one slide, inspect slide 1. For a multi-slide deck, use the slide or slides named by the user; if none is named and the target is unclear, ask instead of reviewing the full deck by default.

Render each relevant slide and run one deep check per relevant slide. If the available renderer exports the whole deck, inspect only the relevant rendered outputs. The slide number is 1-based:

```bash
TARGET_SLIDE=3  # replace with the inspected slide's 1-based number
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/render_pptx.py" \
  "$INPUT_PPTX" --output-dir "$BUILD_DIR/rendered"
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$INPUT_PPTX" --slide "$TARGET_SLIDE" \
  --output "$BUILD_DIR/check-slide-$TARGET_SLIDE.json"
```

If LibreOffice or `pdftoppm` is unavailable, use the host's existing renderer and state which renderer was used; do not select an authoring backend merely to inspect a deck.

Report meaningful content, topology, editability, and package discrepancies. Do not loop, silently repair, or hide uncertainty in invisible objects.

## Compatibility check

Do not add a routine UI check to every task. When a real local PowerPoint installation is accessible, open or export the final authored candidate once only if it contains a compatibility risk—custom/freeform labels, Office Math, dense or mixed-script typography, true vertical or rotated text, a prior repair prompt—or the user explicitly asks for native or cross-platform validation. Check only for a repair prompt, missing content, material reflow or clipping, displaced labels, and broken formulas.

LibreOffice rendering is a server smoke check, not proof of PowerPoint identity. State the actual validation environment and do not create Windows/macOS variants unless a material blocker is reproduced and each named result can be tested there.

For Reconstruction or Repair, return a short note stating the authoring runtime, render and structure checks performed, any risk-triggered native PowerPoint check, and any material approximation or local raster inset. Inspection returns findings only.
