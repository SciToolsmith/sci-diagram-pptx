---
name: sci-diagram-pptx
description: Reconstruct, repair, or inspect user-provided scientific and academic schematic diagrams as visually faithful, native editable PowerPoint (.pptx) files built primarily from PowerPoint shapes, text, and connectors. Use for 复刻、还原、临摹或修复科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots whose meaning is encoded by axes, scales, legends, or data-driven geometry; general business diagrams; organization charts; ordinary deck design; OCR-only extraction; or simple image placement.
---

# SciDiagram PPTX

Rebuild the supplied scientific schematic as a practical, editable PowerPoint. Reproduce the source; do not redesign it unless the user explicitly asks for redesign.

## Preserve what matters

Apply strict fidelity to:

1. visible wording, symbols, formulas, and scientific meaning;
2. nodes, regions, nesting, reading order, and connector topology;
3. native editability of the diagram's main structure;
4. clipping, overlap, missing content, and file integrity.

Treat small differences in antialiasing, kerning, exact color, line weight, corner radius, arrowhead proportions, or local spacing as cosmetic when they do not affect readability or meaning. Do not spend repeated passes chasing pixel identity.

## Confirm the target

Use this skill when meaning is carried primarily by boxes, labels, mathematical expressions, nested regions, arrows, or explicit relationships. Do not use it for bar, line, scatter, violin, box, density, forest, volcano, heatmap, or similar statistical plots.

For a mixed multi-panel figure, reconstruct only the schematic panel explicitly selected by the user. Ask for target selection when the intended panel is unclear. Ask for content confirmation only when an unreadable character, formula, or arrow direction could change meaning; do not interrupt for ordinary styling uncertainty.

Never silently rewrite wording, improve hierarchy, recolor the figure, or reverse a relationship because another design appears clearer.

## Use the presentation runtime

Load and follow the installed `Presentations` skill for every PPTX build, repair, render, or inspection. Apply these reconstruction-specific rules:

- use `@oai/artifact-tool` from a JavaScript ES module for PPTX authoring;
- use Python only for image inspection, cropping, and package checks;
- match the source canvas, density, typography, and orientation instead of applying a generic slide template;
- use native PowerPoint shapes, text, connectors, fills, and strokes for the editable reconstruction;
- do not substitute image generation, Graphviz, imported SVG, outlined text, or source-image tiles for native reconstruction;
- read current Artifact Tool documentation before using unfamiliar APIs, but do not probe standard documented shapes and text again.

Resolve the current workspace dependencies and use the bundled Python executable reported by the workspace. Keep temporary code, renders, and check reports in a task-specific build directory. Put only the final PPTX in the user destination and never overwrite an existing file silently.

## Follow one reconstruction workflow

### 1. Inspect and map the source

Inspect the selected image at original resolution. Identify its major regions, nodes, visible text, connectors, decorative arrows, repeated styles, and isolated raster content.

Create a lightweight object map directly in the build code or its data constants. Record only what is needed to build: stable IDs, text, approximate bounds, object type, and source/target relationships. Do not create a separate Scene Plan, confidence ledger, hash chain, or approval manifest.

When the user explicitly selects a panel from a larger image, crop only that panel without resizing or retouching. The bundled helper is available when exact pixel bounds are useful:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/panel_crop.py" \
  "$PARENT_IMAGE" --bbox <x> <y> <width> <height> \
  --output "$BUILD_DIR/panel.png"
```

### 2. Build with native semantic units

Read [native-object-policy.md](references/native-object-policy.md). Use one coherent native object per meaningful unit whenever practical: a labeled node as one shape with text, a relationship as one connector, and a standard arrow as one shape.

Preserve the source aspect ratio and derive placement from one consistent source-to-slide transform. Keep connectors behind nodes where appropriate and preserve source direction, endpoint, dash style, crossings, and z-order.

Default to two slides using the same canvas:

1. Slide 1 — native editable reconstruction;
2. Slide 2 — the unchanged source image or explicitly selected panel for reference.

Honor an explicit user request for a one-slide deliverable. Slide 1 must never be a full-slide source bitmap, hidden tracing image, or image-tile mosaic.

For formulas, true vertical text, compound freeforms, gradients, or unusual connectors, read [math-and-fonts.md](references/math-and-fonts.md) and, only if support is uncertain, [capability-matrix.md](references/capability-matrix.md). Use one focused capability probe, not a full evidence workflow.

Use a close native approximation without pausing when the difference is cosmetic. Ask first when a fallback changes scientific meaning, formula content, arrow topology, or converts a substantial meaning-bearing region to raster.

### 3. Render and review

Render the exported PPTX through the current `Presentations` workflow. Inspect the reconstruction beside the source at normal size, then enlarge dense text, formulas, and connector crossings.

Fix these blocking defects:

- guessed or incorrect text, symbols, formulas, or arrow directions;
- missing major nodes, regions, labels, or relationships;
- wrong topology, nesting, reading order, or meaning-bearing line style;
- visible clipping, overflow, serious overlap, or off-canvas content;
- flattened or falsely editable main structure;
- corrupt export or missing native content.

If blockers exist, repair them together in one focused pass and render once more. After the second render, continue only for a named blocking defect. Do not rebuild again merely to improve a pixel-difference score.

### 4. Run the lightweight package check

Read [quality-checklist.md](references/quality-checklist.md), then run:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$FINAL_PPTX" --source "$SOURCE_IMAGE" \
  --output "$BUILD_DIR/check.json"
```

The check must confirm that the package opens, contains a usable native reconstruction, avoids whole-page raster simulation, has no macro/OLE/external-media shortcuts, and—when a reference slide exists—uses the supplied source without an additional crop. Fix any reported hard failure. Treat its warnings as prompts for judgment, not automatic rejection.

### 5. Deliver

Deliver when scientific content, topology, native editability, readability, and file integrity are sound. Remaining cosmetic differences are not blockers.

Return the final PPTX and a short note covering:

- that the file was rendered and structurally checked;
- any material native approximation or local raster element;
- any remaining compatibility limitation worth the user's attention.

Do not deliver build scripts, reports, renders, or intermediate files unless requested.

## Handle an existing PPTX naturally

When an editable PPTX is supplied, preserve correct native objects and repair only the mismatches against the authoritative source. Keep the original file untouched and write a new output. When the user asks only for inspection, do not modify the file; render it, run the same lightweight check, and report the meaningful discrepancies.

If a required native behavior is unavailable, explain the exact limitation and the narrowest practical alternative. Do not loop indefinitely or hide uncertainty in invisible objects.
