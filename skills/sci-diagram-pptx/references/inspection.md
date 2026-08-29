# Inspection

Read this file only when an existing PPTX must be diagnosed without modification.

Do not modify the file, select an authoring backend, create `build.mjs`, or create a delivery folder. If the deck has one slide, inspect slide 1. For a multi-slide deck, use the slide or slides named by the user; if none is named and the target is unclear, ask instead of reviewing the full deck by default.

Render each relevant slide and run one deep check per relevant slide. If the renderer exports the whole deck, inspect only the relevant outputs. The slide number is 1-based:

```bash
TARGET_SLIDE=3  # replace with the inspected slide's 1-based number
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/render_pptx.py" \
  "$INPUT_PPTX" --output-dir "$QA_DIR/rendered"
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/check_pptx.py" \
  "$INPUT_PPTX" --slide "$TARGET_SLIDE" \
  --output "$QA_DIR/check-slide-$TARGET_SLIDE.json"
```

If LibreOffice or `pdftoppm` is unavailable, use the host's existing renderer and state which renderer was used. Do not select an authoring backend merely to inspect a deck.

Report meaningful content, topology, editability, and package discrepancies. Do not loop, silently repair, or hide uncertainty in invisible objects.
