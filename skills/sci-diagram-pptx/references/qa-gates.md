# Scientific Diagram QA Gates

Pass every applicable gate on the exported PPTX. Do not validate only the authoring source or intermediate preview.

## Gate 1 — Contract and target

- Confirm that the target is a scientific or academic schematic covered by the skill.
- Reject statistical or data plots, general business diagrams, organization charts, and commercial infographics before implementation.
- For a mixed figure, confirm that the user explicitly selected one schematic panel and record its exact bounds.
- Confirm the active filename, exactly two slides, slide size, and comparison-slide contract.
- Confirm that Slide 1 and Slide 2 refer to the same selected source or panel crop.
- Confirm that the page orientation and aspect ratio match the selected source or panel.
- Resolve every material item in the uncertainty ledger or record explicit user approval.

## Gate 2 — Structural fidelity

- Account for every major region, node, boundary, label, connector, and repeated motif.
- Verify node types, nesting, grouping intent, and front-to-back order.
- Verify connector origin, destination, direction, path style, and arrowhead.
- Verify solid versus dashed semantics and all cross-region relationships.
- Confirm that no important element lies outside the slide canvas.

## Gate 3 — Content and typography

- Compare all visible text against the source at enlarged scale.
- Verify spelling, capitalization, punctuation, abbreviations, and line breaks.
- Verify mathematical symbols, Greek letters, operators, functions, and script positions.
- Inspect Chinese, Latin, and mathematical glyphs for fallback or corruption.
- Fix clipping, overflow, unwanted wrapping, overlap, and baseline drift.

## Gate 4 — Native editability

Inspect the exported deck with Artifact Tool and, when possible, reimport it.

- Confirm that Slide 1 is not one full-slide image.
- Confirm that ordinary text remains editable text.
- Confirm that boxes, fills, borders, lines, arrowheads, and colors remain editable.
- Confirm that formulas retain the exact object type promised to the user.
- Confirm that connectors retain attachment when attachment was promised.
- Reject outlined glyphs, SVG-derived path clouds, per-character fragments, and micro-segment dashes.
- Confirm that object granularity remains practical for later editing.

Treat the machine audit as structural evidence, not a substitute for the final render review. It hard-checks object kind, visible text, bounding boxes, preset geometry, raster use, formula text, and core connector semantics. Theme-resolved fills, strokes, fonts, fine connector routing, and connection-site indices still require full-size visual and manual verification unless the active runtime exposes a proven round-trip API for them.

## Gate 5 — Background and reference slide

- Inspect Slide 1 for full-slide rectangles, full-slide images, hidden source images, and transparent page-sized placeholders; remove them.
- Preserve local fills that genuinely belong to nodes or regions.
- Confirm that Slide 2 uses the original source image or exact selected panel crop without redrawing it.
- Confirm that Slide 2 preserves important content, aspect ratio, centering, and readable scale.
- Confirm that Slide 1 contains no adjacent plot, photograph, table, or other unselected panel from a mixed figure.

## Gate 6 — Visual comparison

Render every exported slide through the bundled render-evidence wrapper and the current Presentations rendering helper. Require the resulting manifest to bind the final PPTX SHA-256 to every rendered-slide SHA-256. Inspect each slide individually at full size; use a montage only for overview.

Compare Slide 1 with the source at three scales:

1. Thumbnail: composition, orientation, color balance, and major omissions
2. Full slide: proportions, alignment, whitespace, hierarchy, and topology
3. Enlarged crops: dense text, equations, arrowheads, corners, dashes, and crossings

Use an overlay or pixel-difference view as diagnostic evidence when useful, but do not let pixel similarity override correct semantics or native editability. Fix all unintended overlap and clipping warnings. Run the current slide-overflow checker on the final PPTX.

Run the bundled overflow wrapper against the current Presentations `slides_test.py`; require its structured report to be `PASS` and bound to the final PPTX SHA-256. After inspecting both final slide renders individually at full size, record a build-bound manual review attestation. It must identify the locked source SHA-256 and final PPTX SHA-256, report `full_size_visual_review: true` and `overflow_check_passed: true`, and cite path-plus-SHA evidence for both renders and the overflow JSON. Never reuse or prefill this attestation.

## Gate 7 — Export integrity

- Open or reimport the final file without repair warnings.
- Verify the final slide count and dimensions after export.
- Verify that embedded source imagery resolves and no temporary paths remain.
- Verify that no placeholders, hidden construction guides, or debug labels remain.
- Verify that only final deliverables appear in the user’s destination.

## Gate 8 — Deviation audit

Match every deviation to explicit user approval. Confirm `approval_source: user-explicit`, approver, timestamp, and evidence are recorded. Confirm that each exception affects only the approved object or region. Remove any broader fallback introduced for convenience.

Do not claim high fidelity, native editability, equation support, or connector attachment beyond what the exported file proves. Return to the relevant policy when a gate fails: [reconstruction-contract.md](reconstruction-contract.md), [native-object-policy.md](native-object-policy.md), or [math-and-fonts.md](math-and-fonts.md).
