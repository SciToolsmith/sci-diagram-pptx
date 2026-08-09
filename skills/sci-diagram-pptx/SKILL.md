---
name: sci-diagram-pptx
description: Reconstruct or repair user-provided scientific and academic schematics as visually faithful, native editable single-slide PowerPoint (.pptx) files with executable build source; also inspect existing PPTX files for meaningful fidelity and editability defects. Use for 复刻、还原、临摹、修复或检查科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots whose meaning is encoded by axes, scales, legends, or data-driven geometry; general business diagrams; organization charts; ordinary deck design; OCR-only extraction; or simple image placement.
---

# SciDiagram PPTX

Rebuild the supplied scientific schematic as a practical, editable PowerPoint. Reproduce the source; do not redesign it unless the user explicitly asks.

## Preserve what matters

Prioritize, in order:

1. visible wording, symbols, formulas, and scientific meaning;
2. nodes, regions, nesting, reading order, and connector topology;
3. native editability of the meaning-bearing structure;
4. readable layout and valid export.

Treat small differences in antialiasing, kerning, exact color, line weight, corner radius, arrowhead proportions, or local spacing as cosmetic when meaning and readability remain sound. Do not chase pixel identity.

Use this skill when meaning is carried mainly by boxes, labels, formulas, nested regions, arrows, or explicit relationships. Do not use it for bar, line, scatter, violin, box, density, forest, volcano, heatmap, or similar statistical plots. For a mixed figure, reconstruct only the schematic panel selected by the user; ask which panel only when the target is unclear. Ask about unreadable content only when a character, formula, or arrow direction could change meaning.

Never silently rewrite wording, simplify topology, recolor the figure, invent a cycle from visual layout, or reverse a relationship because another design appears clearer.

## Select one authoring runtime

Keep one user workflow and choose one host-provided authoring backend before writing `build.mjs`:

1. Honor an explicit `SCI_DIAGRAM_RUNTIME=artifact-tool` or `SCI_DIAGRAM_RUNTIME=pptxgenjs` value.
2. When unset, prefer `artifact-tool` only when the installed OpenAI `Presentations` runtime is available; otherwise select `pptxgenjs` only when its package and server render dependencies are available.
3. Run the bundled runtime probe before authoring when availability is uncertain. Stop with a clear dependency error if the selected backend is unavailable.
4. Do not switch backend after `build.mjs` has been written, and do not place both backends in one delivered script.

Read exactly one runtime reference:

- [runtime-artifact-tool.md](references/runtime-artifact-tool.md) for Codex desktop and the OpenAI `Presentations` runtime;
- [runtime-pptxgenjs.md](references/runtime-pptxgenjs.md) for a host-configured Linux/Node.js deployment.

The reconstruction contract, object semantics, review limits, and delivery are identical. Runtime-specific rendering differences do not justify changing labels, topology, or editability.

## Follow one reconstruction workflow

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

Write `build.mjs` before export and execute that exact file to create the delivered `editable.pptx`. It may import only the selected authoring package and `node:`-prefixed built-ins; keep reconstruction data in the file and do not import local helper modules. Resolve every companion relative to `import.meta.url`. State the selected runtime, tested version, and run command in the header. Write `editable.pptx` beside the script and refuse silent overwrite.

Read other references only when their trigger applies:

- [math-and-fonts.md](references/math-and-fonts.md): formulas, mixed scripts, true vertical text, or dense typography;
- [cross-platform-compatibility.md](references/cross-platform-compatibility.md): custom/freeform labels, fragile typography, complex formulas, a reproduced platform problem, or an explicit cross-platform request;
- [capability-matrix.md](references/capability-matrix.md): only when support for an unusual native feature is uncertain.

Use one focused capability probe when necessary. Use a close native approximation without pausing when the difference is cosmetic. Ask first when a fallback changes scientific meaning, formula content, arrow topology, or converts a substantial meaning-bearing region to raster.

### 3. Review once, then correct once if needed

After the first export, run one render through the selected runtime and one lightweight package check. Compare the render beside the source at normal size; enlarge dense text, formulas, connector crossings, arrowheads, and raster insets only where needed. Confirm every planned edge exists once with the correct endpoint, direction, bidirectionality, label, and route. Inspect each inset at 100% for framing and readable evidence.

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$FINAL_PPTX" --source "$DELIVERY_SOURCE" --build-source "$BUILD_MJS" \
  --require-single-slide \
  --output "$BUILD_DIR/check.json"
```

Collect all blocking defects before editing: incorrect or guessed content; missing, extra, or reversed relationships; wrong nesting or line semantics; unreadable critical insets; clipping, overflow, serious overlap, or off-canvas content; flattened structure; corrupt export; or a package hard failure.

If blockers exist, repair them together in one focused pass, then render and run the same package check once more. Continue after that only for a named blocking defect. Warnings require judgment, not automatic rejection; cosmetic differences are not blockers.

### 4. Use native PowerPoint only for compatibility risk

Do not add a routine UI check to every task. When a real local PowerPoint installation is accessible, open or export the final candidate once only if the file contains a compatibility risk—custom/freeform labels, Office Math, dense or mixed-script typography, true vertical or rotated text, a prior repair prompt—or the user explicitly asks for native or cross-platform validation. Check only for a repair prompt, missing content, material reflow or clipping, displaced labels, and broken formulas.

LibreOffice rendering is a server smoke check, not proof of PowerPoint identity. State the actual validation environment and do not create Windows/macOS variants unless a material blocker is reproduced and each named result can be tested there.

### 5. Deliver the executable folder

For image-to-PPTX reconstruction, return:

```text
<diagram-name>_editable/
├── source.<original-extension>
├── editable.pptx
└── build.mjs
```

The source is the unchanged upload, the PPTX contains exactly one native editable reconstruction slide, and `build.mjs` is the exact executed program. The script is reproducible in a compatible declared runtime; it is not a bundled package manager environment.

Validate the staged folder before handoff: compare source bytes, confirm the staged script was executed, and reject machine-local paths, secrets, or imports of unshipped helpers. Keep checks, renders, probes, temporary crops, caches, and intermediate exports internal. Do not add a README, manifest, preview, ZIP, `node_modules`, or package-manager files unless requested. Include `assets/` only for separately supplied companions required by the build. For several independent images, create one folder per target.

Return the folder with a short note stating the authoring runtime, render and structure checks performed, any risk-triggered native PowerPoint check, and any material approximation or local raster inset.

## Repair or inspect an existing PPTX

For repair, keep the original untouched and write a new output. Prefer a self-contained rebuild. If the script depends on the original deck, include a byte-for-byte copy such as `input-original.pptx` and resolve it relatively; include a supplied reference image too. Do not describe a dependency-bearing repair folder as the standard three-file image bundle.

For inspection only, do not modify the file or create a delivery folder. Render the relevant slides with the available host runtime, then run the checker without source or single-slide assertions:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$INPUT_PPTX" --output "$BUILD_DIR/check.json"
```

Report meaningful content, topology, editability, and package discrepancies. If a required native behavior is unavailable, state the exact limitation and narrowest practical alternative; do not loop or hide uncertainty in invisible objects.
