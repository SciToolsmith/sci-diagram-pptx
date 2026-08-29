---
name: sci-diagram-pptx
description: Reconstruct or repair user-provided scientific and academic schematics as visually faithful, native editable PowerPoint (.pptx) files with executable build source; also inspect existing PPTX files for meaningful fidelity and editability defects. Use for 复刻、还原、临摹、修复或检查科研框架图、技术路线图、科研流程图、机制图、算法流程图、科研系统结构图、学术概念模型、结构化科研信息图，以及包含数学符号或公式的学术图示. Do not use for quantitative data visualizations or statistical plots whose meaning is encoded by axes, scales, legends, or data-driven geometry; general business diagrams; organization charts; ordinary deck design; OCR-only extraction; or simple image placement.
---

# SciDiagram PPTX

Reproduce the supplied scientific schematic; do not redesign it unless the user explicitly asks.

## Preserve what matters

Prioritize, in order:

1. visible wording, symbols, formulas, and scientific meaning;
2. nodes, regions, nesting, reading order, and connector topology;
3. native editability of the meaning-bearing structure;
4. readable layout and valid export.

Treat small differences in antialiasing, kerning, exact color, line weight, corner radius, arrowhead proportions, or local spacing as cosmetic when meaning and readability remain sound. Do not chase pixel identity.

Use this skill when meaning is carried mainly by boxes, labels, formulas, nested regions, arrows, or explicit relationships. Do not use it for charts whose evidence is encoded by axes, scales, legends, or data-driven geometry. Split a mixed figure by origin and reconstruct only the selected schematic panel. Ask about unreadable content only when a character, formula, or arrow direction could change meaning.

Never silently rewrite wording, simplify topology, recolor the figure, invent a cycle from visual layout, or reverse a relationship because another design appears clearer.

## Choose exactly one route

- **Reconstruction**: an image or selected image panel becomes a new single-slide editable PPTX. Read [reconstruction.md](references/reconstruction.md).
- **Repair**: named defects in an existing PPTX are corrected without changing unrelated slides. Read [repair.md](references/repair.md).
- **Inspection**: named slides in an existing PPTX are diagnosed without modification. Read [inspection.md](references/inspection.md).

Read only the selected route reference. Do not load the other route contracts. If a request mixes routes, complete the named deliverable and report separate findings briefly.

## Select one runtime only when authoring

Reconstruction and Repair choose one host-provided backend before writing `build.mjs`:

1. Honor an explicit `SCI_DIAGRAM_RUNTIME=artifact-tool` or `SCI_DIAGRAM_RUNTIME=pptxgenjs` value.
2. When unset for Reconstruction, prefer `artifact-tool` when the installed OpenAI `Presentations` runtime is available; otherwise use `pptxgenjs` only when its locked package and render dependencies are ready.
3. Repair requires a backend that can import and preserve the existing deck. Use `artifact-tool` when available. PptxGenJS cannot perform Repair; stop or obtain approval to change the task to standalone Reconstruction.
4. Probe only when availability is uncertain. Do not repeat a successful probe for the same backend, version, task directory, and host unless the environment changes or a real invocation fails.
5. Do not switch backend after `build.mjs` exists or mix both authoring packages in one script.

Read exactly one runtime reference:

- [runtime-artifact-tool.md](references/runtime-artifact-tool.md) for the desktop OpenAI `Presentations` runtime;
- [runtime-pptxgenjs.md](references/runtime-pptxgenjs.md) for the locked portable backend.

Inspection selects no authoring backend.

## Shared construction invariants

- Record relationships from the source, not from proximity or circular placement. A line label is not a node. Preserve origin, destination, direction, dash meaning, crossings, nesting, and z-order.
- Give every visible element exactly one owner: **native** or **raster**. Text or arrows retained inside an inset must not also be redrawn natively. If a native overlay is essential, exclude or cover the raster copy in the initial ownership map rather than discovering duplicates through repeated renders.
- Keep photographs, microscopy, experimental results, model screenshots, and complex composite mini-figures as replaceable local raster insets when native reconstruction would be dishonest. Keep their semantic frame and external annotations native.
- Never use a full-slide source bitmap, hidden tracing image, imported path cloud, or image-tile mosaic on the editable slide.
- Preserve source aspect ratio and use one source-to-slide transform. Keep topology and object data explicit in the executed build source.
- Ask before a fallback changes scientific meaning, formula content, arrow topology, or converts a substantial meaning-bearing region to raster. Use a close native approximation without pausing when the difference is cosmetic.

## Bound the work

For Reconstruction and Repair, the default budget is two authored passes at most:

1. Build `pass-01`, then run the selected render and the package checker. Run those two read-only checks concurrently when the host supports it.
2. Combine checker hard failures with named visual blockers. Warnings and cosmetic differences do not authorize another pass.
3. If no blocker remains, deliver the accepted pass.
4. If blockers remain, collect all of them before editing and make one grouped correction in `pass-02`.
5. The second render/check is terminal. Deliver or report the named blocker. A third authored pass requires an explicit user request.

Use `--pass-index 1` and `--pass-index 2` with the checker. Its `decision.next_action` is `deliver`, `repair_once`, or `stop_with_blocker`. Do not create `final`, `approved`, `release`, or similarly renamed extra passes. Do not use a pixel-difference percentage as a substitute for checking wording, formulas, topology, and visible clipping.

Inspection is read-only and never starts a repair loop.

## Read optional references only when triggered

- [math-and-fonts.md](references/math-and-fonts.md): mathematical notation, mixed scripts, true vertical text, or dense typography;
- [cross-platform-compatibility.md](references/cross-platform-compatibility.md): only an explicit cross-platform request or a reproduced application-specific discrepancy;
- [capability-matrix.md](references/capability-matrix.md): only an unusual native feature whose support is genuinely uncertain.

Use at most one focused capability probe for one unresolved feature. Reuse a result already established for the same backend and version.

## Portability and handoff

Use portable construction plus the normal render and structure check as the validation boundary. A successful render proves only the selected renderer opened and exported the file; it does not prove pixel identity across applications. Do not create platform variants by default.

For Reconstruction or Repair, return a short note stating the authoring runtime, accepted pass, checks performed, and any material approximation or local raster inset. Inspection returns findings only.
