# Artifact Tool Runtime

Read this file only after selecting `artifact-tool` for Reconstruction or Repair. Do not load the PptxGenJS runtime reference in the same task.

## Required Presentations context

Load the installed OpenAI `Presentations` Skill and follow its hard requirements. Read only the Presentations resources required by the active route:

- **Every Artifact Tool authoring task**: read its required `style_guidelines.md`, `artifact_tool_docs/API_QUICK_START.md`, and `artifact_tool_docs/api/API_DOCS.md` for runtime setup, API usage, export, and render instructions.
- **Reconstruction from an image**: treat the supplied image and the SciDiagram contract as explicit visual direction. Do not load template-following, Codex Grid, built-in templates, or unrelated deck-design resources.
- **Repair of a user-provided PPTX**: also read the Presentations `references/template-following.md` and perform the source-deck inspection it requires so masters, layouts, and unrelated slides are preserved. Do not mix in Codex Grid or another template.

SciDiagram overrides general deck-design defaults only where they conflict with faithful schematic reconstruction: do not create a narrative arc, redesign the source, substitute Graphviz/ImageGen, enforce generic deck font minima, or place machine-absolute paths in delivered `build.mjs`. Keep Presentations requirements for safe workspace setup, source preservation, Artifact Tool implementation, rendering, overlap/clipping checks, and final delivery.

## Build

1. Initialize a task-local Artifact Tool workspace with the setup helper supplied by `Presentations`.
2. Write one self-contained `build.mjs` that imports `@oai/artifact-tool` plus only `node:fs`, `node:fs/promises`, `node:path`, or `node:url`, and writes the adjacent `editable.pptx`.
3. For Reconstruction, read the adjacent `source.*`. For Repair, read `input-original.pptx` or another adjacent companion only when the executable repair genuinely depends on it.
4. Keep nodes, directed edges, labels, and raster insets explicit; never infer a cycle or arrow direction from visual placement.
5. Export and render according to the selected Reconstruction or Repair contract, then run the shared checker for the target slide or slides.

Use tested connector helpers with a documented source-to-target convention. In the current Artifact Tool exporter, `slide.shapes.connect(from, to, ...)` maps `head` to the source/start and `tail` to the target/end. Therefore a forward `from -> to` connector uses `tail`; a bidirectional connector uses `head` and `tail`. Treat edge labels as labels, not nodes. Create endpoint nodes first, then call `connect`; the tested exporter places resulting connectors behind their endpoints. Verify the named arrow ends in the focused render instead of reversing `from` and `to` to compensate visually.

## Scope

Artifact Tool is the desktop-preferred backend and may use capabilities verified by its current documentation. For uncertain equations, true vertical text, compound freeforms, or unusual connector routing, run one minimal capability probe rather than expanding the main workflow.

The runtime changes implementation only. It does not change faithful topology, native editability, local replaceable raster insets, or the route-specific delivery contract. Reconstruction uses the default three-file bundle; separately supplied build dependencies may add `assets/`, and Repair includes only the adjacent dependencies its executable build requires.
