# Native Object Policy

## Rebuild semantic units

Represent each meaningful unit with the simplest PowerPoint-native object that preserves how a user would edit it.

| Source unit | Preferred PowerPoint object |
| --- | --- |
| Labeled node | One shape containing native text |
| Standalone label | One text box with coherent text runs |
| Rectangle, rounded box, ellipse, diamond, arrow | Matching preset shape |
| Simple irregular polygon | One native custom shape when supported |
| Relationship | One connector or native line with arrowhead |
| Boundary or divider | One native outline or line with dash style |
| Intrinsically raster inset | One local image object at the correct layer |

Keep text as text and standard arrows as single shapes or connectors. Do not fragment labels into characters, arrows into shaft-and-head pieces, or dashed borders into many short lines.

## Preserve topology and practical editability

Preserve connector origin, destination, direction, solid or dashed meaning, arrowhead, major crossings, reading order, and nesting. Put edges behind nodes where this improves readability. Name important objects when the authoring API supports names.

Group only coherent modules when grouping helps later editing. Avoid a page-wide group and deep nesting. Optimize for direct editing and recognizable structure rather than maximizing object count.

## Limit raster content

Slide 1 must not be a full-slide source image, hidden tracing image, imported SVG path cloud, or image-tile mosaic. Allow a raster object only when the source element is intrinsically raster, such as a microscopy thumbnail or texture. Keep its labels, frame, and relationships native.

Ask before rasterizing a substantial meaning-bearing region. A small cosmetic approximation may be implemented directly and disclosed at delivery.

Do not add a page-sized background shape merely to imitate a white source canvas. Preserve local fills that belong to the diagram itself.
