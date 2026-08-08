# Math and Fonts

Read this file when the source contains mathematical notation, mixed scripts, true vertical text, or dense typography.

## Preserve content before appearance

Transcribe visible labels and formulas before formatting. Preserve language, capitalization, punctuation, operators, Greek letters, superscripts, subscripts, and meaningful line breaks. Ask the user only when an unreadable character could change meaning; do not pause for an uncertain font family or minor spacing detail.

Keep ordinary Chinese, English, numerals, and symbols as editable PowerPoint text. Put a node label inside a standard preset shape when practical; overlay a standard text box on custom geometry or freeforms unless their text rectangle has been verified. Avoid one object per character, outlined glyphs, SVG text, and formula screenshots when native text is adequate.

## Match mathematical complexity

- Use native text and structured runs with the native baseline property for short expressions, variables, Greek letters, and simple superscripts or subscripts. Do not prefer Unicode super/subscript lookalikes when a structured baseline run is available.
- Prefer a genuine Office Math object for fractions, matrices, roots, stacked limits, multi-level scripts, and other complex notation when the current runtime can create, export, and preserve it reliably.
- When complex notation cannot remain both faithful and editable, ask the user to choose between editable linear notation and a local formula image. Do not guess or silently alter the formula.

Prioritize semantic correctness, editability, readability, placement, and then pixel-level similarity—in that order.

## Choose fonts pragmatically

Use the exact font when it is installed and clearly identifiable. Otherwise choose a visually close, stable font with complete glyph coverage and dependable availability in the intended Office environments. Prefer one portable font choice over silent per-platform substitutions. Use embedding only when the font license and toolchain support it, and confirm that embedding actually occurred before claiming it.

Write the chosen Latin, East Asian, and complex-script font mappings explicitly when the API exposes them. Set font size, paragraph alignment, line spacing, text-box margins, and meaningful line breaks during generation. Size the final text box before export and leave practical width and height headroom—roughly 8%–12% is a useful compact-label heuristic, not a universal pass/fail threshold. Prefer enlarging the text region or making a small source-faithful line-break adjustment before reducing font size.

Do not depend on PowerPoint to resize text or geometry when the file opens. Emit stable final bounds and typography. If the authoring runtime requires AutoFit metadata, treat it only as a safety hint and verify that the exported layout is already correct without platform-specific reflow.

Render once at full slide size and inspect dense regions for missing glyphs, fallback jumps, broken wrapping, clipped accents, and baseline drift. Treat small antialiasing, kerning, or font-shape differences as non-blocking cosmetic differences. Use [cross-platform-compatibility.md](cross-platform-compatibility.md) for portable output and native smoke-check policy.
