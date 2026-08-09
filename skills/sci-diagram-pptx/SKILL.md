---
name: sci-diagram-pptx
description: Reconstruct or repair user-provided scientific and academic schematics as visually faithful, native editable single-slide PowerPoint (.pptx) files with reproducible build source; also inspect existing PPTX files for meaningful fidelity and editability defects. Use for 复刻、还原、临摹、修复或检查科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots whose meaning is encoded by axes, scales, legends, or data-driven geometry; general business diagrams; organization charts; ordinary deck design; OCR-only extraction; or simple image placement.
---

# SciDiagram PPTX

Rebuild the supplied scientific schematic as a practical, editable, portable PowerPoint. Reproduce the source; do not redesign it unless the user explicitly asks.

## Preserve what matters

Prioritize, in order:

1. visible wording, symbols, formulas, and scientific meaning;
2. nodes, regions, nesting, reading order, and connector topology;
3. native editability of the meaning-bearing structure;
4. readable layout and valid export.

Treat small differences in antialiasing, kerning, exact color, line weight, corner radius, arrowhead proportions, or local spacing as cosmetic when meaning and readability remain sound. Do not chase pixel identity.

Use this skill when meaning is carried mainly by boxes, labels, formulas, nested regions, arrows, or explicit relationships. Do not use it for bar, line, scatter, violin, box, density, forest, volcano, heatmap, or similar statistical plots. For a mixed figure, reconstruct only the schematic panel selected by the user; ask which panel only when the target is unclear. Ask about unreadable content only when a character, formula, or arrow direction could change meaning.

Never silently rewrite wording, simplify topology, recolor the figure, or reverse a relationship because another design appears clearer.

## Apply the reconstruction contract

Load the installed `Presentations` skill for its current Artifact Tool runtime, API documentation, rendering, and package utilities. When its general deck-design rules conflict with this skill, this reconstruction contract is authoritative:

- The supplied scientific figure is the visual and structural specification. Do not apply Codex Grid, a narrative arc, title-slide conventions, generic deck templates, or a redesigned information hierarchy.
- Build the diagram with programmatic native PowerPoint shapes, text, lines, and connectors. Do not substitute Graphviz, image search, ImageGen, imported SVG, outlined text, source-image tiles, or other generated artwork for the reconstruction.
- Match the source typography and density. Generic 50/35/24/16 pt deck minimums do not apply; smaller source-faithful text is allowed when it remains readable.
- Use absolute paths in the temporary execution environment when required, but the delivered `build.mjs` must resolve every companion relative to `import.meta.url` and contain no machine-specific absolute path.
- For repair or inspection, review only the slides placed in scope by the user; do not invoke template-following or inspect unrelated slides unless the task covers the whole deck.

Use `@oai/artifact-tool` from a JavaScript ES module for authoring. Use Python only for image inspection, cropping, and package checks. Resolve workspace dependencies and use the bundled Python executable reported by the workspace. Keep temporary crops, renders, probes, and check output in a task-specific build directory. Never overwrite an existing user file or delivery folder silently.

## Follow one reconstruction workflow

### 1. Inspect and map

Inspect the selected image at original resolution. Identify major regions, nodes, visible text, connectors, repeated styles, and intrinsic raster elements. Keep a lightweight object map in `build.mjs` or its data constants: stable IDs, text, approximate bounds, object type, and source/target relationships are enough. Do not create a separate Scene Plan, confidence ledger, hash chain, approval manifest, or QA report.

When the user explicitly selects a panel from a larger image, crop only that panel without resizing or retouching:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/panel_crop.py" \
  "$SOURCE_IMAGE" --bbox <x> <y> <width> <height> \
  --output "$BUILD_DIR/reference-panel.png"
```

Keep the exact upload as `SOURCE_IMAGE`. Copy it byte-for-byte into the staged folder as `source.<original-extension>` and call that copy `DELIVERY_SOURCE`. Use the original as `REFERENCE_IMAGE` unless a panel was selected; then use the temporary crop for comparison and record its pixel bounds in `build.mjs`, while still delivering the unchanged full upload.

### 2. Build native semantic units

Use the simplest coherent native object for each meaningful unit: a preset shape may own its label; a standalone label is one text box; a relationship is one connector or line; an intrinsic raster inset is one local image. Keep text as text and preserve connector origin, destination, direction, dash meaning, crossings, nesting, and z-order.

Treat custom geometry and freeforms as visual plates with overlaid standard text boxes unless the exported geometry contains a usable text rectangle and a focused render verifies it. Do not fragment labels into individual characters, arrows into shaft-and-head pieces, or dashed borders into many short lines. Do not use a full-slide source bitmap, hidden tracing image, imported path cloud, or image-tile mosaic on the editable slide.

Preserve the source aspect ratio and derive placement from one consistent source-to-slide transform. Create exactly one slide: the native editable reconstruction. Do not add a source-reference or hidden tracing slide.

Write `build.mjs` before export and execute that exact file to create the delivered `editable.pptx`. Keep the object map and constants in it whenever practical. It may import only `@oai/artifact-tool` and `node:`-prefixed Node built-ins; keep all reconstruction data in the file and do not import local helper modules. A short header must state that the Artifact Tool runtime is required and give the run command. Write `editable.pptx` beside the script. If it already exists, require an explicit documented overwrite flag or stop with a clear error; never replace it silently.

Normal preset-shape, text-box, and connector reconstructions need no additional reference file. Read references only when their trigger applies:

- [math-and-fonts.md](references/math-and-fonts.md): formulas, mixed scripts, true vertical text, or dense typography;
- [cross-platform-compatibility.md](references/cross-platform-compatibility.md): custom/freeform labels, fragile typography, complex formulas, a reproduced platform problem, or an explicit cross-platform request;
- [capability-matrix.md](references/capability-matrix.md): only when support for an unusual native feature is uncertain.

Use one focused capability probe when necessary. Use a close native approximation without pausing when the difference is cosmetic. Ask first when a fallback changes scientific meaning, formula content, arrow topology, or converts a substantial meaning-bearing region to raster.

### 3. Review once, then correct once if needed

After the first export, complete one review round containing both checks. Run them in parallel when the available tools make that practical; otherwise run them back-to-back before editing:

1. one render through the current `Presentations` workflow, compared beside the source at normal size with dense text, formulas, and crossings enlarged as needed; and
2. one lightweight package check:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$FINAL_PPTX" --source "$DELIVERY_SOURCE" --build-source "$BUILD_MJS" \
  --require-single-slide \
  --output "$BUILD_DIR/check.json"
```

Collect all blocking defects before editing:

- incorrect or guessed text, symbols, formulas, or arrow directions;
- missing major nodes, regions, labels, or relationships;
- wrong topology, nesting, reading order, or meaning-bearing line style;
- clipping, overflow, serious overlap, or off-canvas content;
- flattened or falsely editable main structure;
- corrupt export, missing native content, or a package hard failure.

If blockers exist, repair them together in one focused pass, then render and run the same package check once more. After that second review, continue only for a named blocking defect. Warnings require judgment, not automatic rejection; cosmetic differences are not blockers.

### 4. Use native PowerPoint only for compatibility risk

Do not add a routine UI check to every task. When a real local PowerPoint installation is accessible, open or export the final candidate once only if the file contains a compatibility risk—custom/freeform labels, Office Math, dense or mixed-script typography, true vertical or rotated text, a prior repair prompt—or the user explicitly asks for native or cross-platform validation. Check only for a repair prompt, missing content, material text reflow or clipping, displaced labels, and broken formulas.

If PowerPoint is unavailable, or no risk trigger applies, do not simulate or claim this check. State the actual validation scope briefly at delivery. Do not create Windows/macOS variants unless a material blocker is reproduced in real PowerPoint on each named environment.

### 5. Deliver the reproducible folder

For image-to-PPTX reconstruction, return:

```text
<diagram-name>_editable/
├── source.<original-extension>
├── editable.pptx
└── build.mjs
```

- `source.<original-extension>` is the byte-for-byte uploaded image.
- `editable.pptx` contains exactly one native editable reconstruction slide.
- `build.mjs` is the actual executed program, uses portable relative companion paths, and documents its Artifact Tool runtime requirement.

Validate the staged folder before handoff: compare source bytes, confirm the staged script is the executed script, and reject machine-local paths, secrets, or imports of unshipped helpers. Keep `check.json`, renders, probes, temporary crops, caches, and intermediate exports internal. Do not add a README, manifest, preview, ZIP, `node_modules`, or package-manager files unless requested. For several independent images, create one folder per target rather than a multi-slide deck.

Return the folder with a short note stating that it was rendered and structurally checked, whether a risk-triggered native PowerPoint check ran and where, and any material approximation or compatibility limitation.

## Repair or inspect an existing PPTX

For repair, keep the original PPTX untouched, preserve correct native objects, and write a new output. The delivered build must be reproducible:

- Prefer a self-contained `build.mjs` reconstruction when practical.
- If `build.mjs` opens or depends on the original deck, include a byte-for-byte copy such as `input-original.pptx` in the delivery folder and resolve it relatively. If a reference image is also supplied, include it too. Do not describe this dependency-bearing folder as the standard three-file image bundle.
- If no reference image exists, do not fabricate one; deliver the repaired PPTX, actual build source, and any required original-deck dependency, and disclose the narrower comparison basis.

Apply single-slide and `--source` flags only when the actual repair contract supports them.

For inspection-only work, do not modify the file or create a delivery folder. Render the relevant slides, then run the checker without source or single-slide assertions:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$INPUT_PPTX" --output "$BUILD_DIR/check.json"
```

Report meaningful content, topology, editability, and package discrepancies. If a required native behavior is unavailable, state the exact limitation and narrowest practical alternative; do not loop or hide uncertainty in invisible objects.
