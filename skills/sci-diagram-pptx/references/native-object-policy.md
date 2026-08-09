# Native Object Policy

## Rebuild semantic units

Represent each meaningful unit with the simplest PowerPoint-native object that preserves how a user would edit it.

| Source unit | Preferred PowerPoint object |
| --- | --- |
| Labeled preset node | One preset shape containing native text |
| Standalone label | One text box with coherent text runs |
| Rectangle, rounded box, ellipse, diamond, arrow | Matching preset shape |
| Custom geometry or freeform with a label | Custom/freeform visual plate plus one standard text box, grouped when useful |
| Custom geometry with a verified text rectangle | One native custom shape containing text when the text rectangle exports and renders correctly |
| Relationship | One connector or native line with arrowhead |
| Boundary or divider | One native outline or line with dash style |
| Intrinsically raster inset | One local image object at the correct layer |

Keep text as text and standard arrows as single shapes or connectors. Do not fragment labels into characters, arrows into shaft-and-head pieces, or dashed borders into many short lines.

Preset rectangles, rounded rectangles, ellipses, diamonds, and standard arrows may own their labels because their text regions are stable. Do not assume the bounding box of an irregular polygon is a safe text region. Use custom geometry and freeforms as visual plates and overlay an ordinary PowerPoint text box by default. Group the plate and text box only when it makes later editing clearer.

Allow a custom/freeform object to own text only after its exported geometry contains an explicit, usable text rectangle and a focused render confirms placement and clipping. The visual path alone is not a text rectangle. If that proof is absent, keep the text separate even when the authoring library accepts a text property.

## Preserve topology and practical editability

Preserve connector origin, destination, direction, solid or dashed meaning, arrowhead, major crossings, reading order, and nesting. Put edges behind nodes where this improves readability. Name important objects when the authoring API supports names.

Group only coherent modules when grouping helps later editing. Avoid a page-wide group and deep nesting. Optimize for direct editing and recognizable structure rather than maximizing object count.

## Limit raster content

The single editable slide must not be a full-slide source image, hidden tracing image, imported SVG path cloud, or image-tile mosaic. Keep the unchanged source outside the PPTX in the delivery folder. Allow a raster object only when the source element is intrinsically raster, such as a microscopy thumbnail or texture. Keep its labels, frame, and relationships native.

Ask before rasterizing a substantial meaning-bearing region. A small cosmetic approximation may be implemented directly and disclosed at delivery.

Do not add a page-sized background shape merely to imitate a white source canvas. Preserve local fills that belong to the diagram itself.
