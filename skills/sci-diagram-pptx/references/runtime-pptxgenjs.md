# PptxGenJS Portable Runtime

Read this file only after the host has selected `pptxgenjs` for the current task. Do not load the Artifact Tool runtime reference in the same task.

## Select and probe

Use this route when the host explicitly sets `SCI_DIAGRAM_RUNTIME=pptxgenjs`, or when Artifact Tool is unavailable and the portable dependency probe reports ready.

Run the read-only probe before authoring:

```bash
node "$SCI_DIAGRAM_SKILL_DIR/scripts/probe_runtime.mjs" --runtime pptxgenjs
```

The route requires Node.js 20+, PptxGenJS 4.0.x, Python, LibreOffice/soffice, and `pdftoppm`. If a required dependency is missing, stop with the probe's JSON result. Do not install software during a user task and do not switch backend after creating `build.mjs`.

## Author from the portable template

Copy `assets/pptxgenjs-build-template.mjs` into the delivery folder as `build.mjs`, then replace its example object data. Keep the final file self-contained and preserve these explicit collections:

- `NODES`: stable IDs, standard shape types, labels, source-coordinate boxes, and styles.
- `EDGES`: stable IDs, `from`, `to`, direction, and an exact point path. The array is the topology; never infer loops or relations from layout.
- `INSETS`: source file, source crop box, destination frame, role, and replacement-friendly object name.

Use `connectOneWay`, `connectBothWays`, or `connectExactPath`. In PptxGenJS 4.0.1, `beginArrowType` exports as OOXML `a:headEnd` at the source/start, while `endArrowType` exports as `a:tailEnd` at the target/end. Therefore `from -> to` uses the final segment's `endArrowType`; a bidirectional relation also uses the first segment's `beginArrowType`. Labels use their own text boxes and never become graph nodes.

Keep source geometry in pixels and map it once into slide inches. Set the source dimensions and slide dimensions explicitly. Add edges first, then raster insets, then node plates and labels.

## Local raster insets

For a photo, microscopy image, result thumbnail, model screenshot, or complex embedded subfigure, keep the semantic frame and annotations native while inserting one cropped image object:

```javascript
{
  id: "replaceable-image-01",
  source: "source.png",
  sourceBox: { x: 820, y: 130, w: 180, h: 110 },
  frameBox: { x: 820, y: 130, w: 180, h: 110 },
  role: "contextual"
}
```

PptxGenJS embeds the source image once and stores a visible `srcRect` crop on the picture. Do not create temporary crop files for delivery. Keep the default bundle as `source.*`, `editable.pptx`, and `build.mjs`. Add `assets/` only when the user supplied independent higher-resolution files required to rebuild the slide.

Do not enlarge a low-resolution crop and call it publication quality. For a scientifically critical unreadable inset, request the original vector, PDF, data, or high-resolution image.

## Render and check

Build once, then render with an isolated LibreOffice profile:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/render_pptx.py" \
  editable.pptx --output-dir qa-render
```

Run the shared structural checker after rendering. LibreOffice proves that the package can be opened and exported on the server; it does not prove pixel identity with PowerPoint on Windows or macOS.

On a web service, treat generated `build.mjs` as untrusted task code. The host—not this Skill—must run it in an isolated per-task directory with network, process, time, memory, and filesystem limits; never expose an interactive shell or server credentials to the task.

## Phase-one boundary

This portable route supports standard rectangles, rounded rectangles, ellipses, diamonds, explicit text boxes, straight/polyline arrows, and local raster insets. Do not silently use it for Office Math, true vertical text, custom freeforms, deep groups, gradients, or complex attached connectors. Stop or obtain approval for the smallest meaning-preserving alternative when those features are essential.
