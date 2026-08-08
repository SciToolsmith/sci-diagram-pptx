# Math and Fonts

Read this file when the source contains mathematical notation, mixed scripts, true vertical text, or dense typography.

## Preserve content before appearance

Transcribe visible labels and formulas before formatting. Preserve language, capitalization, punctuation, operators, Greek letters, superscripts, subscripts, and meaningful line breaks. Ask the user only when an unreadable character could change meaning; do not pause for an uncertain font family or minor spacing detail.

Keep ordinary Chinese, English, numerals, and symbols as editable PowerPoint text. Put a node label inside its owning shape when practical. Avoid one object per character, outlined glyphs, SVG text, and formula screenshots when native text is adequate.

## Match mathematical complexity

- Use native text and structured runs for short expressions, variables, Greek letters, and simple super/subscripts.
- Use a genuine Office Math object only when the current runtime can create and preserve it reliably.
- When complex notation cannot remain both faithful and editable, ask the user to choose between editable linear notation and a local formula image. Do not guess or silently alter the formula.

Prioritize semantic correctness, editability, readability, placement, and then pixel-level similarity—in that order.

## Choose fonts pragmatically

Use the exact font when it is installed and clearly identifiable. Otherwise choose a visually close, stable font with complete glyph coverage. Font substitution is acceptable when it does not create missing glyphs, different meaning, clipping, or seriously altered wrapping.

Render once at full slide size and inspect dense regions for missing glyphs, fallback jumps, broken wrapping, clipped accents, and baseline drift. Treat small antialiasing, kerning, or font-shape differences as non-blocking cosmetic differences.
