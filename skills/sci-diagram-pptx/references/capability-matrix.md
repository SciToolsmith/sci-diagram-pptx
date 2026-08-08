# Sci-Diagram Capability Matrix

## Route all PPTX authoring through Artifact Tool

After confirming that sci-diagram-pptx applies, load the current Presentations skill before implementation. Use @oai/artifact-tool from a JavaScript ES module for every local PPTX build. Initialize its prescribed temporary workspace, read API_QUICK_START.md and api/API_DOCS.md, and export through PresentationFile.exportPptx.

Do not use python-pptx, the retired Python artifact_tool API, SVG-to-PowerPoint conversion, or a second authoring library. Do not pin instructions to a versioned cache path; resolve paths from the loaded Presentations skill.

## Probe before promising

Classify each requested feature before full reconstruction:

- Direct: documented, constructible, editable, and preserved after export.
- Conditional: plausible or importable, but require a focused round-trip probe.
- Approved fallback: unavailable natively but replaceable after user consent.
- Blocked: no faithful route and no approved fallback.

For every conditional feature:

1. Search the current Artifact Tool docs and runtime help for the exact API.
2. Build the smallest representative object in the temporary workspace.
3. Export it to PPTX.
4. Reimport or inspect the exported file and render it.
5. Confirm object type, text editability, geometry, and visible result.
6. Record the result and use only the tested route.

Do not treat “no export error” as proof of semantic preservation.

## Apply the matrix

Treat capability and scope as separate gates. Artifact Tool support for charts does not authorize reconstruction of Python, R, MATLAB, or statistical data plots, commercial diagrams, or organization charts with this skill.

| Feature | Preferred Artifact Tool route | Initial class | If the probe fails |
| --- | --- | --- | --- |
| Custom slide size | Presentation.create with source-matched dimensions | Direct | Stop if dimensions cannot survive export |
| Preset nodes and arrows | slide.shapes.add with preset geometry | Direct | Try one native custom shape |
| Polygonal freeforms | One custom shape with validated paths | Direct for documented line paths | Approximate only with approval when geometry changes |
| Curved or compound freeforms | Documented custom geometry, if available | Conditional | Use a close preset or raster exception only with approval |
| Native text boxes and shape text | Shape text bodies and structured runs | Direct | Stop if export flattens text |
| Bold, italic, underline, color | Structured text runs | Direct | Use supported run formatting only |
| Superscript, subscript, character spacing | Exact documented property or proven round trip | Conditional | Offer editable linear or Unicode alternatives |
| True vertical text direction | Exact documented property or proven round trip | Conditional | Offer rotated editable text and disclose the difference |
| Native lines and dashes | LineConfig and connector endpoints | Direct | Use a single native custom shape if suitable |
| Arrowhead endpoint mapping | Probe head/tail export on one directed connector | Conditional | Use only the verified endpoint mapping; stop if direction cannot be preserved |
| Attached connectors and elbows | slide.shapes.connect with validated anchors | Direct | Use a free-positioned native line and disclose lost attachment |
| Z-order | Native bring-to-front/send-to-back operations | Direct | Recreate in correct creation order |
| Grouping | Documented group API plus export test | Conditional | Leave semantic objects ungrouped |
| Office Math equation | Documented equation API or preserved native template object | Conditional | Follow the math fallback approval flow |
| Intrinsic raster inset inside the selected schematic | Embedded image bytes with native crop and placement | Direct | Preserve aspect ratio; do not redraw as fake vectors |
| Whole-diagram bitmap | None on the reconstruction slide | Blocked by policy | No fallback within this skill; route elsewhere |
| Imported SVG or outlined text | None as the primary reconstruction route | Blocked by policy | Rebuild with native objects or route elsewhere |

Treat “Direct” as version-sensitive. Re-probe when the installed tool version, export path, or requested property differs from the validated case.

## Request degradation approval

Present one compact decision request containing:

- The affected object IDs or source regions
- The desired native behavior
- The capability that failed and the evidence
- Two or three feasible alternatives
- The visual, semantic, and editability cost of each alternative
- A recommended narrow fallback

Wait for explicit approval before implementing a material degradation. Preserve unaffected elements under the original native-object contract. Record the approved exception for final QA.

Record approval provenance in the scene plan: `approval_source: user-explicit`, `approved_by`, `approved_at`, and a concise `evidence` excerpt or interaction reference. A model-authored `approved: true` without this provenance is not approval.
