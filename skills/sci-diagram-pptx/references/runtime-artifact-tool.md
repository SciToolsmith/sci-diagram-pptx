# Artifact Tool Runtime

Read this file only after the host has selected `artifact-tool` for the current task. Do not load the PptxGenJS runtime reference in the same task.

## Select this runtime

Use this route when the installed OpenAI `Presentations` Skill and `@oai/artifact-tool` workspace runtime are available. An explicit host value `SCI_DIAGRAM_RUNTIME=artifact-tool` takes precedence over automatic selection. If the requested dependency is unavailable, stop before authoring and report it; do not switch an existing `build.mjs` to another backend.

## Build

1. Load the installed `Presentations` Skill for runtime setup, Artifact Tool API, export, and render instructions. For this source-faithful reconstruction, the SciDiagram contract overrides general deck-design defaults such as narrative redesign, layout grids, Graphviz/ImageGen substitution, generic minimum font sizes, or machine-absolute paths in the delivered script.
2. Read its Artifact Tool quick start and API documentation before writing code.
3. Initialize a task-local Artifact Tool workspace with the setup helper supplied by `Presentations`.
4. Write one self-contained `build.mjs` that imports only `@oai/artifact-tool` and `node:` built-ins, reads the adjacent source image, and writes the adjacent `editable.pptx`.
5. Keep the scientific object map explicit. Record nodes, directed edges, labels, and raster insets as data; never infer a cycle or arrow direction from visual placement.
6. Export one slide, render once with the `Presentations` helpers, and run the shared PPTX checker.

Use tested connector helpers with a documented source-to-target convention. In the current Artifact Tool exporter, `slide.shapes.connect(from, to, ...)` maps `head` to the source/start and `tail` to the target/end. Therefore a forward `from -> to` connector uses `tail`; a bidirectional connector uses `head` and `tail`. Treat edge labels as labels, not nodes. Create endpoint nodes first, then call `connect`; the tested exporter places resulting connectors behind their endpoints. Verify the named arrow ends in the focused render instead of reversing `from` and `to` to compensate visually.

## Scope

Artifact Tool remains the desktop-preferred backend and may use capabilities verified by its current documentation. For uncertain equations, true vertical text, compound freeforms, or unusual connector routing, run one minimal capability probe rather than expanding the main workflow.

The runtime changes implementation only. It does not change the shared requirements for one editable slide, faithful topology, local replaceable raster insets, or the three-file delivery bundle.
