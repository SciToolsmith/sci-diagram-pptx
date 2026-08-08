# Practical Quality Checklist

Run this checklist on the exported PPTX. The goal is a correct, editable, visibly faithful reconstruction—not a machine proof of pixel identity.

## Blocking defects

Fix these before delivery:

- unreadable text, formulas, or arrow directions were guessed;
- a major region, node, label, or relationship is missing;
- connector direction, topology, nesting, or reading order is wrong;
- text is clipped, meaning-bearing objects overlap, or content lies outside the canvas;
- Slide 1 is flattened, tiled, or primarily composed of the source bitmap;
- the PPTX cannot open, render, or retain its expected native objects;
- a material approximation or raster fallback changes meaning without user confirmation.

## Non-blocking differences

Do not hold delivery for small differences in font rasterization, kerning, exact color, line weight, corner radius, arrowhead proportions, or local spacing when the diagram remains readable and scientifically correct. Mention only material differences that a user may want to adjust.

## Review sequence

1. Render the exported presentation once through the current Presentations workflow.
2. Inspect the whole reconstruction beside the source at normal reading size.
3. Enlarge dense text, formulas, and connector crossings.
4. Run `scripts/check_pptx.py` for basic package and editability checks.
5. If a blocking defect exists, repair all visible blockers in one focused pass, render again, and rerun the lightweight check once.
6. When a real local PowerPoint installation is accessible, open or export the final candidate once as a native smoke check.
7. Deliver when no blocker remains. Do not continue iterating only to improve a pixel score.

After the second render, make further edits only for a named blocking defect. If critical content remains unreadable or a required native behavior is unsupported, ask the user rather than looping indefinitely.

## Portability checks

- Block `customGeom/freeform + text + no explicit text rectangle` in a generated deliverable because the same label can be placed safely in a standard text box. Inventory other custom geometry as a warning rather than rejecting it merely for being custom.
- Inspect AutoFit dependence, inadequate text headroom, missing glyphs, unexpected font fallback, and fragile character fragments.
- Keep the optional native smoke check to one final open or export. Do not add separate Windows/macOS review modes or a multi-round approval workflow.
- Deliver one portable PPTX unless [cross-platform-compatibility.md](cross-platform-compatibility.md) justifies separately validated derivatives.
