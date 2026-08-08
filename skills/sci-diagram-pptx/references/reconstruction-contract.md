# Scientific Diagram Reconstruction Contract

## Core promise

Use sci-diagram-pptx to reconstruct a selected scientific or academic schematic as a PowerPoint-native, editable PPTX. Treat 复刻, 还原, 临摹, “rebuild,” and “recreate” as fidelity requests unless the user explicitly asks for redesign.

Include research frameworks, technical routes, scientific process diagrams, mechanism diagrams, algorithm flowcharts, system structures or architectures, concept diagrams, scientific infographics, and academic figures containing symbols or formulas.

Exclude Python, R, MATLAB, or other data and statistical visualizations, including bar, line, area, scatter, bubble, histogram, box, violin, density, heat-map, correlation, volcano, forest, and similar plots. Exclude general business processes, organization charts, commercial roadmaps, and non-academic infographics. Route those requests to an appropriate charting, coding, or presentation workflow instead of stretching this contract.

## Isolate a schematic panel

Treat a multi-panel image as mixed when it combines an in-scope schematic with plots, microscopy, photographs, tables, or other panels. Process it only after the user explicitly identifies the schematic panel by label, description, or bounds.

Reconstruct only the selected panel. Do not trace, approximate, or embed adjacent excluded panels on Slide 1. Use the exact selected panel crop as the default Slide 2 reference. Stop and ask for panel selection when the requested bounds are unclear or when the user asks to reconstruct the entire mixed figure.

When the workflow materializes a panel crop, use `scripts/panel_crop.py` and preserve its manifest. Record the parent SHA-256 and displayed pixel dimensions, zero-based panel bounds, crop label, manifest SHA-256, and `panel_selection` provenance (`selection_source: user-explicit`, selector, time, and evidence). Validate the scene plan against that manifest so the parent bytes, crop pixels, bounds, and user selection remain linked.

## Resolve the target

1. Require the authoritative source image or exact selected panel crop; return `NEEDS_SOURCE` for source-less Audit or Repair requests.
2. Identify the exact source image and confirm that the target is an in-scope scientific schematic.
3. Use the only uploaded image when it contains one unambiguous in-scope figure.
4. Require explicit panel selection for every mixed figure, even when only one image was uploaded.
5. Follow the user’s explicit image and panel selection when several candidates exist.
6. Stop and ask when several targets or panel bounds remain plausible.
7. Use the same selected image or panel crop for reconstruction and visual reference.
8. Inspect orientation, regions, text blocks, object types, connectors, visual hierarchy, and dense details before authoring.

## Preserve the source

Reproduce content, structure, geometry, relationships, and styling. Do not beautify, modernize, simplify, normalize, or reinterpret the source merely because another design looks cleaner.

Prioritize fidelity in this order:

1. Correct target and complete content
2. Correct wording, labels, symbols, and mathematical meaning
3. Global composition and page aspect ratio
4. Module positions, dimensions, and whitespace
5. Connector endpoints, direction, and line semantics
6. Shape geometry and hierarchy
7. Color relationships, borders, dashes, corners, and z-order
8. Fine spacing and pixel-level appearance

Preserve asymmetry, density, unusual colors, and imperfect alignment when they are genuine source characteristics. Correct only production defects introduced during reconstruction.

## Use the fixed deliverable contract

Deliver exactly two slides:

- Slide 1: the editable native reconstruction.
- Slide 2: the unredrawn selected source or exact schematic-panel crop, preserving all target-panel content and centered at its original aspect ratio.

Use the same slide size for both slides. Match the selected source or panel orientation and aspect ratio instead of forcing 16:9. Let an explicit user filename and destination win; otherwise use a non-colliding `可编辑复现版.pptx`. Do not remove the reference slide or change the two-slide architecture under this skill; route a materially different deliverable to a general presentation workflow.

## Control inference and degradation

Infer only low-risk visual details whose alternatives do not change meaning, topology, or editability. Never invent illegible wording, mathematical characters, arrow direction, node identity, or missing relationships.

Maintain an uncertainty ledger with the affected region, observed evidence, confidence, alternatives, and consequence of a wrong choice. Resolve material ambiguity before committing the affected objects.

When a required feature cannot be reproduced natively:

1. Verify the limitation through the capability probe in [capability-matrix.md](capability-matrix.md).
2. Describe the exact affected elements and fidelity dimension.
3. Offer feasible alternatives with editability and visual tradeoffs.
4. Recommend the narrowest acceptable fallback.
5. Obtain explicit approval before applying it.
6. Limit approval to the named elements; do not generalize it to the whole slide.

Do not silently rasterize, outline text, fragment objects, substitute uncertain content, or change structure. Never manufacture an approval: record who approved it, when, and the explicit user evidence. Stop and ask when the unresolved choice is material and no approved fallback exists.

## Complete the contract

Model Slide 1 according to [native-object-policy.md](native-object-policy.md), handle formulas and typography according to [math-and-fonts.md](math-and-fonts.md), and pass every gate in [qa-gates.md](qa-gates.md) before delivery.
