# Repair

Read this file only when named defects in an existing PPTX must be corrected while unrelated content is preserved.

Identify the exact slide number or numbers and defect before authoring; ask only when the target is unclear. Keep the original PPTX untouched and preserve unrelated slides, masters, layouts, notes, and order.

Repair requires a backend that can import and preserve the deck. The bundled PptxGenJS route cannot perform Repair. Stop with that limitation or, with user approval, change the task to standalone Reconstruction.

Prefer a self-contained rebuild. If `build.mjs` must read the original deck, include a byte-for-byte `input-original.pptx` beside it and resolve that dependency relative to `import.meta.url`. Include a supplied reference image only when the repair or its build actually depends on it.

For `pass-01`, render the repaired result and deep-check each repaired slide separately. Do not pass `--source` or `--require-single-slide` merely because a visual reference was used or the original deck has one slide:

```bash
TARGET_SLIDE=3  # replace with the repaired slide's 1-based number
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$PASS_DIR/editable.pptx" --build-source "$PASS_DIR/build.mjs" \
  --slide "$TARGET_SLIDE" --pass-index 1 \
  --output "$QA_DIR/check-pass-01-slide-$TARGET_SLIDE.json"
```

Collect hard failures and named visual blockers across all repaired slides. Warnings alone do not authorize another pass. If blockers exist, make one grouped correction in `pass-02` and rerun only affected slide checks with `--pass-index 2`. The second check is terminal.

Return a repair folder containing `editable.pptx`, the exact executed `build.mjs`, and only adjacent input dependencies required to rerun it. Do not describe it as the standard three-file Reconstruction bundle. Move or copy the accepted pass byte-for-byte; do not repeat unchanged checks after staging.
