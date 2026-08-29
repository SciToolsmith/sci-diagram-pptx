# PptxGenJS Portable Runtime

Read this file only after the host has selected `pptxgenjs` for the current task. Do not load the Artifact Tool runtime reference in the same task.

## Select and probe

Use this route when the host explicitly sets `SCI_DIAGRAM_RUNTIME=pptxgenjs`, or when Artifact Tool is unavailable and the portable dependency probe reports ready.

Run the read-only probe before authoring:

```bash
node "$SCI_DIAGRAM_SKILL_DIR/scripts/probe_runtime.mjs" \
  --runtime pptxgenjs --task-dir "$BUILD_DIR"
```

`--task-dir` must be the directory that will contain and execute `build.mjs`; it defaults to the current directory. The probe resolves the bare `pptxgenjs` import from that exact location and requires a detected 4.0.x release. A package visible only from the Skill repository or the caller's unrelated current directory does not make the task ready.

Run this probe once per task directory. Do not repeat it after a successful result unless the directory, dependency tree, or runtime version changes or a real invocation fails.

The route also requires Node.js 20+, Python 3.10+, LibreOffice/soffice, and `pdftoppm`. If a required dependency is missing, stop with the probe's JSON result. Do not install software during a user task and do not switch backend after creating `build.mjs`.

## Host deployment contract

Install the repository's locked Node dependencies once in a host-owned runtime root, then create task directories below that root so Node's normal ESM lookup can reach the ancestor `node_modules`:

```text
<runtime-root>/
├── package.json + package-lock.json + node_modules/   # host-owned, never delivered
└── tasks/<task-id>/                                   # BUILD_DIR
    ├── source.png
    └── build.mjs
```

Run `npm ci --ignore-scripts --no-audit --no-fund` only while provisioning the host runtime root. Do not copy `node_modules`, `package.json`, or `package-lock.json` into a task or delivery folder. Before authoring each task, run the probe with that task's `--task-dir`; after the host dependency is reachable, execute `node "$BUILD_DIR/build.mjs"` from the same directory. The final delivery remains only `source.*`, `editable.pptx`, and `build.mjs`.

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
  role: "contextual",
  sourceOrientation: "normalized"
}
```

PptxGenJS embeds the source image once and stores a visible `srcRect` crop on the picture. Do not create temporary crop files for delivery. Keep the default bundle as `source.*`, `editable.pptx`, and `build.mjs`. Add `assets/` only when the user supplied independent higher-resolution files required to rebuild the slide.

The locked 4.0.1 template path is covered by the bundled crop test. For a normalized source using that exact route, do not add a task-specific crop probe.

`srcRect` cropping is not reliable for a JPEG whose active EXIF orientation is not `1`: PptxGenJS preserves the metadata, while server renderers may crop the raw pixel matrix without applying the display orientation. Inspect orientation before declaring `sourceOrientation: "normalized"`. If the selected source has active EXIF rotation or mirroring, stop this portable route and request an orientation-normalized replacement source; do not transform displayed crop coordinates by guesswork, create an undeclared companion, or switch backend mid-task.

Do not enlarge a low-resolution crop and call it publication quality. For a scientifically critical unreadable inset, request the original vector, PDF, data, or high-resolution image.

## Render and check

Build once, then render with an isolated LibreOffice profile:

```bash
"$SCI_DIAGRAM_PYTHON" "$SCI_DIAGRAM_SKILL_DIR/scripts/render_pptx.py" \
  editable.pptx --output-dir qa-render
```

Run the shared structural checker after rendering. LibreOffice proves that the package can be opened and exported on the server; it does not prove pixel identity with PowerPoint on Windows or macOS.

On a web service, treat generated `build.mjs` as untrusted task code. The host—not this Skill—must run it in an isolated per-task directory with network, process, time, memory, and filesystem limits; never expose an interactive shell or server credentials to the task.

## Existing PPTX repair boundary

PptxGenJS does not natively import and edit an existing PPTX. Stop before authoring when the request is Repair and report that this backend cannot preserve the original deck. Only after the user explicitly approves changing the task to standalone slide Reconstruction may you render the untouched input as a visual reference and generate a new single-slide file under the Reconstruction contract. Never describe that result as Repair, copy the old file, patch only its ZIP parts, claim that unseen objects, masters, layouts, notes, animations, or groups were preserved, or switch backend after `build.mjs` exists.

## Phase-one boundary

This portable route supports standard rectangles, rounded rectangles, ellipses, diamonds, explicit text boxes, straight/polyline arrows, and local raster insets. Do not silently use it for Office Math, true vertical text, custom freeforms, deep groups, gradients, or complex attached connectors. Stop or obtain approval for the smallest meaning-preserving alternative when those features are essential.
