# Complex Element Guide

Read this file only when the source contains a feature whose PowerPoint behavior is genuinely uncertain: complex equations, true vertical text, compound freeforms, unusual connector routing, deep groups, gradients, or embedded scientific imagery. Read the selected runtime reference first. Do not probe ordinary rectangles, text boxes, preset arrows, lines, or connectors already marked as tested for that backend.

## Use one focused probe

When a capability could affect scientific meaning or editability:

1. Search the selected runtime documentation for the exact property.
2. Build the smallest representative object.
3. Export and render it once.
4. Inspect the exported object type and visible result.
5. Reuse the proven route in the reconstruction.

Stop probing after the smallest test answers the question. A full-slide prototype is unnecessary.

## Choose the least disruptive implementation

| Source feature | Preferred implementation | Acceptable practical fallback |
| --- | --- | --- |
| Standard nodes and block arrows | Native preset shape | Closest native preset with a disclosed cosmetic difference |
| Irregular but simple polygon | One native custom/freeform visual plate; use a separate standard text box for its label | Closest editable preset when meaning is unchanged |
| Curved or compound freeform | Supported custom geometry as a visual plate; use a separate standard text box for its label | Editable approximation after confirming that topology is unchanged |
| Native text and mixed formatting | Shape text or structured runs | Installed close font if all glyphs and line breaks remain correct |
| Superscript, subscript, vertical text | Native baseline/structured runs for scripts; proven native property for vertical text | Editable Unicode, linear text, or rotation when meaning remains clear |
| Complex equation | Native Office Math when reliably supported | Ask before using linear text or a local formula image |
| Attached connector | Tested runtime helper with explicit source, target, and direction | Free-positioned native line when attachment is not essential |
| Intrinsic raster inset | One replaceable picture object with recorded source crop | Preserve it as raster, keep its frame and annotation native, and request a better asset if critical content is unreadable |
| Whole diagram bitmap or image tiles | Never on the editable reconstruction | No fallback within this skill |

Use a close native approximation without interrupting the user when the difference is purely cosmetic. Ask first when a fallback changes a formula, label, arrow direction, topology, scientific interpretation, or converts a substantial meaning-bearing region to raster.

Mention material approximations and local raster elements in the delivery note. Do not create manifests, approval ledgers, or capability reports unless the user explicitly requests them.
