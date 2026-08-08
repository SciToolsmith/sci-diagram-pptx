# Scientific Diagram Native Object Policy

## Build a clean editable canvas

Keep Slide 1 transparent at page level. Do not add a full-slide white rectangle, colored plate, background image, hidden source image, or transparent placeholder shape. Preserve local white or colored fills only when they belong to the diagram itself.

Test the canvas conceptually by copying all reconstructed objects onto a colored slide: empty space inside and around the diagram must reveal the destination background.

## Map semantics to native objects

Represent each meaningful unit with the simplest PowerPoint-native object that preserves its behavior:

| Source unit | Preferred PowerPoint object |
| --- | --- |
| Labeled box or node | One shape containing native text |
| Standalone label | One text box with structured text runs |
| Rectangle, rounded box, ellipse, diamond, arrow | Matching preset shape |
| Irregular polygon | One native custom shape when supported |
| Relationship | One connector or line with native arrowheads |
| Boundary or divider | One line or shape outline with native dash style |
| Intrinsically raster content | One cropped image object at the correct layer |

Keep ordinary text as text. Model only the selected scientific schematic; do not use chart objects or diagram shapes to imitate excluded statistical plots. Do not flatten diagram structure into an image.

## Preserve useful granularity

Use one semantic unit per object whenever practical. Put a node’s text inside its shape instead of creating a redundant overlay text box. Keep a complete label, simple formula, arrow, border, or dashed line intact.

Avoid these fragmentation patterns:

- One text object per character or word
- Separate objects for a basic arrow shaft and head
- Dozens of short segments for one dashed line
- Multiple paths for a shape available as one preset
- Imported SVG paths used as the primary reconstruction method
- Text converted to outlines, glyph paths, or vector fragments

Optimize for fidelity, direct editing, and maintainability—not the smallest or largest possible object count.

## Preserve topology and layering

Create relationship edges before nodes when practical so labels and nodes remain readable. Attach connectors to the correct source and target boundaries. Preserve direction, endpoint, crossing, elbow or curve intent, solid or dashed semantics, arrowhead type, and line weight.

Keep connectors away from labels unless the source explicitly overlaps them. Reproduce intentional occlusion and z-order. Inspect arrowheads after export because line caps and previews can vary across renderers.

Use native outline and dash settings. Do not simulate standard borders or dashes with micro-shapes. Use native freeform or custom shapes only when a preset cannot match the geometry and the runtime preserves editability.

## Group conservatively

Group only semantically coherent modules when grouping is supported and helps later movement. Keep individual boxes, labels, and connectors directly editable. Avoid a single page-wide group and deeply nested groups.

Name important objects by role or source region when the authoring API supports names. Use stable names to simplify inspection and repair; do not expose internal labels on the slide.

## Limit raster content

Allow raster objects on Slide 1 only for an intrinsically raster inset that belongs inside the selected schematic panel, such as a microscopy thumbnail or source texture that cannot meaningfully become PowerPoint geometry. Keep surrounding labels, frames, and relationships native.

Do not copy adjacent panels from a mixed figure onto Slide 1. Treat heat maps and other data plots as out of scope even when they are already rasterized.

Treat rasterizing any editable candidate as a degradation. Follow the approval process in [capability-matrix.md](capability-matrix.md). Never use an approved local raster exception as permission to rasterize the entire visual.
