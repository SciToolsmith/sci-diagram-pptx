# Reconstruction

Read this file only when an image or selected image panel must become a new single-slide editable PPTX.

## Inspect and assign ownership

Inspect the target at original resolution. Keep a lightweight source-coordinate map inside `build.mjs` with these collections when applicable:

- `NODES`: stable ID, visible text, type, source-derived bounds, and style;
- `EDGES`: stable ID, `from`, `to`, direction, exact route, line semantics, and optional label;
- `INSETS`: stable ID, source pixel box, slide frame, semantic role, and the content intentionally owned by the raster crop.

Assign each visible label, formula, arrow, mini-plot, or illustration to exactly one collection. If an inset owns internal text or arrows, do not duplicate them in `NODES` or `EDGES`. If surrounding titles, borders, captions, or annotations will be native, choose source boxes that exclude them. When exclusion is impossible, plan one deliberate cover and replacement before the first full build; do not accumulate masking patches across passes.

Represent reciprocal relationships explicitly as bidirectional. Never replace a pair, branch, or open path with a guessed loop.

Classify photographs, microscopy, experimental results, model screenshots, and complex composite mini-figures as intrinsic raster insets. Do not OCR-guess, AI-upscale, sharpen, or redraw unreadable scientific evidence. Request a better asset only when a critical inset is unreadable at its intended size.

When the user selects one panel from a larger image, crop only that panel for temporary comparison without resizing or retouching:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/panel_crop.py" \
  "$SOURCE_IMAGE" --bbox <x> <y> <width> <height> \
  --output "$QA_DIR/reference-panel.png"
```

Keep the exact upload as `SOURCE_IMAGE`. Deliver its byte-for-byte copy as `source.<original-extension>` even when a temporary panel crop is the visual reference. Record selected-panel pixel bounds in `build.mjs`.

## Build native semantic units

Use the simplest coherent native object for each meaningful unit: a preset shape may own its label; a standalone label is one text box; a relationship is one connector or line; an intrinsic raster inset is one replaceable picture object. Use the selected runtime's tested connector convention rather than ambiguous raw arrow-end calls.

Treat custom geometry and freeforms as visual plates with overlaid standard text boxes unless the exported geometry has a usable text rectangle and a focused render verifies it. Do not fragment labels into individual characters, arrows into shaft-and-head pieces, or dashed borders into many short lines.

Crop insets from `source.*` through picture crop metadata and record their source boxes. Do not deliver separate crop files. If two or more insets use the same source and crop export for the exact backend/version has not already been established on the current host, test one representative crop before the full build. Reuse that result for every inset; never probe each crop. Skip this probe when the locked runtime template or a successful current-host result already establishes the behavior.

Create exactly one slide: the editable reconstruction. Do not add a reference or hidden tracing slide.

Write `build.mjs` before export and execute that exact file. Import exactly one selected authoring package plus only the safe built-ins accepted by the checker: `node:fs`, `node:fs/promises`, `node:path`, and `node:url`. Keep reconstruction data in the file; do not use dynamic imports, `require`, `createRequire`, or local helper modules. Resolve every companion relative to `import.meta.url`. State the runtime, tested version, and run command in the header. Refuse silent overwrite.

## Run the bounded passes

Create a task-local `pass-01/` containing the source copy and `build.mjs`; let the executed script create its adjacent `editable.pptx`. Put renders and reports in a sibling QA directory so an accepted pass already contains only delivery files.

After the first export, start one render and one package check. Run them concurrently when possible:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$PASS_DIR/editable.pptx" --source "$PASS_DIR/source.<ext>" \
  --build-source "$PASS_DIR/build.mjs" --require-single-slide --slide 1 \
  --pass-index 1 --output "$QA_DIR/check-pass-01.json"
```

Use the checker's compact stdout for control. Read only `decision`, `hard_failures`, and material `warnings` from the JSON; do not dump the full report into the conversation.

Compare the render beside the source once at normal size, then enlarge only dense text, formulas, connector crossings, arrowheads, and critical insets. A visual blocker is incorrect or guessed content; a missing, extra, or reversed relationship; wrong nesting or line semantics; an unreadable critical inset; clipping, serious overlap, or off-canvas content; a flattened structure; a corrupt export; or a package hard failure. Small spacing, antialiasing, font-shape, color, or arrowhead differences are cosmetic.

If blockers exist, list them all before editing. Create `pass-02/`, apply one grouped correction, and rerun the same checks with `--pass-index 2`. Stop after that result even if a named blocker remains; continue only when the user explicitly requests another attempt.

## Deliver the accepted pass

```text
<diagram-name>_editable/
├── source.<original-extension>
├── editable.pptx
└── build.mjs
```

The PPTX contains exactly one native editable reconstruction slide, and `build.mjs` is the exact program that generated it. If the accepted pass is moved or copied byte-for-byte into the delivery location, compare bytes and do not re-execute, re-render, or re-check unchanged files. Remove internal QA artifacts before handoff.

Reject machine-local paths, secrets, and imports of unshipped helpers. Do not add a README, manifest, preview, ZIP, `node_modules`, or package-manager files unless requested. Include `assets/` only for separately supplied companions required by the build. For several independent images, create one delivery folder per target.
