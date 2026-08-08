# Cross-Platform Compatibility

Use this policy for every reconstruction. Target one portable, editable PPTX that remains readable and scientifically correct across PowerPoint environments; do not promise pixel identity.

## Build one portable file

Generate one portable PPTX by default. Keep content, object IDs, topology, and layout in one canonical build rather than maintaining Windows and macOS branches. Do not infer compatibility from the generation host or label an untested file for another platform.

Create platform-specific derivatives only when:

1. a material blocker is reproduced in real PowerPoint on the affected environments;
2. portable object, typography, and formula rules cannot remove it;
3. variants come from the same canonical object data and differ only where necessary; and
4. each named variant is validated in real PowerPoint on its named environment.

If exact appearance matters more than editability, offer a fixed PDF preview alongside the single editable PPTX instead of automatically creating two editable files.

## Prefer stable native structures

Use preset shapes, standard text boxes, native connectors, native lines, and coherent text runs wherever they express the source correctly. Preset shapes with stable text regions may contain labels. Treat custom geometry and freeforms as visual plates with overlaid standard text boxes unless the exported geometry has an explicit text rectangle and one focused render verifies it. Align the plate and text with shared bounds or a shallow group.

The checker must fail a generated deliverable that contains `customGeom + text + no text rectangle`; the same label can be represented safely as an ordinary text box over the custom visual plate. Inventory other custom geometry as a compatibility warning rather than rejecting it merely for being custom. Do not treat a successful export from the authoring library as proof that PowerPoint applications will place that text consistently.

## Freeze typography during generation

Resolve layout before export: choose explicit fonts with full glyph coverage and realistic availability; set sizes, paragraph properties, margins, and meaningful line breaks; size the final text regions with practical headroom; then render and make source-faithful box, margin, line-break, or minimal font-size adjustments before exporting stable bounds.

Roughly 8%–12% spare width and height is a useful heuristic for compact labels when the source geometry permits it. It is not a universal threshold: long lines, CJK text, mathematical notation, rotated labels, and narrow shapes may need different margins. Judge the actual rendered result and preserve visual fidelity.

Do not depend on open-time AutoFit to choose a font size, resize a shape, or rewrap text. If AutoFit metadata cannot be avoided, the layout must already fit without it and the metadata must not be the only protection against clipping.

## Keep formulas semantic

Use native baseline formatting in structured text runs for simple superscripts and subscripts. Prefer Office Math for fractions, roots, matrices, stacked limits, and other complex equations when the runtime can create and preserve it reliably. Avoid Unicode super/subscript lookalikes as the default because coverage and metrics vary by font.

When a complex formula cannot remain both faithful and editable, ask before using editable linear notation or a local formula image. A platform-specific version is not a substitute for resolving ambiguous scientific notation.

## Perform one final native smoke check

After normal render and package checks pass, use a real local Microsoft PowerPoint installation once when accessible: open or export the final candidate and check for a repair prompt, missing objects, material text reflow or clipping, displaced custom/freeform labels, and broken formulas. Record the actual environment. If PowerPoint is unavailable, disclose that validation was limited to the documented renderer and package checks. Never imply that an untested platform was validated; this smoke check is not a platform matrix or renewed multi-round review.
