---
name: sci-diagram-pptx
description: Reconstruct, repair, or inspect user-provided scientific and academic schematic diagrams as visually faithful, native editable single-slide PowerPoint (.pptx) files, delivered with the unchanged source image and executed build.mjs. Use for 复刻、还原、临摹或修复科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots whose meaning is encoded by axes, scales, legends, or data-driven geometry; general business diagrams; organization charts; ordinary deck design; OCR-only extraction; or simple image placement.
---

# SciDiagram PPTX

Rebuild the supplied scientific schematic as a practical, editable, portable PowerPoint. Reproduce the source; do not redesign it unless the user explicitly asks for redesign.

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

Resolve the current workspace dependencies and use the bundled Python executable reported by the workspace. Keep probes, renders, check reports, and temporary crops in a task-specific build directory. Stage the final deliverable as one folder and never overwrite an existing file or folder silently.

## Follow one reconstruction workflow

### 1. Inspect and map the source

Inspect the selected image at original resolution. Identify its major regions, nodes, visible text, connectors, decorative arrows, repeated styles, and isolated raster content.

Create a lightweight object map directly in the build code or its data constants. Record only what is needed to build: stable IDs, text, approximate bounds, object type, and source/target relationships. Do not create a separate Scene Plan, confidence ledger, hash chain, or approval manifest.

When the user explicitly selects a panel from a larger image, crop only that panel without resizing or retouching. The bundled helper is available when exact pixel bounds are useful:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/panel_crop.py" \
  "$SOURCE_IMAGE" --bbox <x> <y> <width> <height> \
  --output "$BUILD_DIR/reference-panel.png"
```

Keep the exact uploaded file as `SOURCE_IMAGE`. Copy it byte-for-byte into the staged delivery folder as `source.<original-extension>` and refer to that copy as `DELIVERY_SOURCE`. Use `SOURCE_IMAGE` as `REFERENCE_IMAGE` when no crop is needed. After an explicit crop, use the selected panel as `REFERENCE_IMAGE` for reconstruction and visual comparison, but keep the unchanged full upload as `DELIVERY_SOURCE`. Record the selected pixel bounds in `build.mjs`; do not deliver the temporary crop by default.

### 2. Build with native semantic units

Read [native-object-policy.md](references/native-object-policy.md) and [cross-platform-compatibility.md](references/cross-platform-compatibility.md). Use one coherent native object per meaningful unit whenever practical. Standard preset shapes may contain their labels. Treat custom geometry and freeforms as visual plates with overlaid standard text boxes unless a usable text rectangle has been explicitly defined and verified.

Preserve the source aspect ratio and derive placement from one consistent source-to-slide transform. Keep connectors behind nodes where appropriate and preserve source direction, endpoint, dash style, crossings, and z-order.

Create exactly one slide: the native editable reconstruction. Do not add a source-reference slide, hidden tracing slide, or hidden source image. The editable slide must never be a full-slide source bitmap or image-tile mosaic.

Write the actual authoring program as `build.mjs` before exporting. It must be the exact program executed to create the delivered `editable.pptx`, not a later summary. Keep the object map and required constants in that file whenever practical. Resolve companion files relative to `import.meta.url`, write `editable.pptx` beside the script, and do not include machine-specific absolute paths, temporary directories, credentials, or secrets. A short header comment may state the `@oai/artifact-tool` runtime requirement and run command; do not add a package manager bundle by default.

For formulas, true vertical text, compound freeforms, gradients, or unusual connectors, read [math-and-fonts.md](references/math-and-fonts.md) and, only if support is uncertain, [capability-matrix.md](references/capability-matrix.md). Size and wrap text during generation with explicit fonts, line breaks, margins, and reasonable headroom; do not depend on PowerPoint changing the layout through open-time AutoFit. Use one focused capability probe, not a full evidence workflow.

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
  "$FINAL_PPTX" --source "$DELIVERY_SOURCE" --require-single-slide \
  --output "$BUILD_DIR/check.json"
```

The check must confirm that the package opens, contains exactly one slide with a usable native reconstruction, avoids whole-page raster simulation, has no macro/OLE/external-media shortcuts, and recognizes the delivered source as an external companion rather than an embedded reference page. Fix reported hard failures together and rerun this lightweight check once on the repaired file. Treat warnings as prompts for judgment, not automatic rejection.

### 5. Run one native smoke check when available

After render review and package checking pass, open or export the final candidate once with a real local PowerPoint installation when one is accessible. Check only for a repair prompt, missing content, material text reflow or clipping, displaced custom geometry labels, and broken formulas. This is a final compatibility smoke check, not a Windows/macOS mode or a multi-round visual audit.

If native PowerPoint is unavailable, do not simulate or claim this check. Record the actual render, package, and platform validation scope in the delivery note.

### 6. Deliver

Deliver when scientific content, topology, native editability, readability, and file integrity are sound. Remaining cosmetic differences are not blockers.

For an image-to-PPTX reconstruction, return one folder with this stable minimal layout:

```text
<diagram-name>_editable/
├── source.<original-extension>
├── editable.pptx
└── build.mjs
```

- `source.<original-extension>` is a byte-for-byte copy of the user's uploaded image; renaming is allowed only to make the bundle portable.
- `editable.pptx` contains exactly one native editable reconstruction slide.
- `build.mjs` is the actual executed source that generated that PPTX and uses relative, portable paths.

Stage and validate the complete folder before moving it to the user destination. Confirm that `DELIVERY_SOURCE` has the same bytes as `SOURCE_IMAGE`, that the staged `build.mjs` was the program actually executed, and that it contains no machine-local absolute path or secret. Keep `check.json`, renders, probes, temporary crops, caches, and intermediate exports internal. Do not add a README, manifest, preview, ZIP, `node_modules`, or package-manager files unless the user explicitly asks for them. If several independent source images are requested, make one three-file folder per target rather than silently combining them into a multi-slide deck.

Return the folder and a short note covering:

- that the file was rendered and structurally checked;
- whether the one native PowerPoint smoke check was performed and on which actual environment;
- any material native approximation or local raster element;
- any remaining compatibility limitation worth the user's attention.

The bundle contains one portable PPTX by default. Derive platform-specific copies only when a real cross-platform blocker remains after the portable construction rules and each copy can be validated in its named target PowerPoint environment. Do not create or label Windows/macOS versions from operating-system assumptions alone.

## Handle an existing PPTX naturally

When an editable PPTX is supplied, preserve correct native objects and repair only the mismatches against the authoritative source. Keep the original file untouched and write a new output. If a reference image is also supplied, use the same three-file bundle and make `build.mjs` the actual repair/rebuild program. If no image exists, do not fabricate one; deliver the repaired PPTX and its actual build source together and disclose that no source-image companion was available. When the user asks only for inspection, do not modify the file or create a delivery folder; render it, run the same lightweight check, and report the meaningful discrepancies.

If a required native behavior is unavailable, explain the exact limitation and the narrowest practical alternative. Do not loop indefinitely or hide uncertainty in invisible objects.
