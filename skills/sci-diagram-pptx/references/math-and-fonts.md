# Scientific Diagram Math and Fonts

## Transcribe before formatting

Separate recognition from layout. Transcribe every visible label and formula first, then verify ambiguous characters against nearby notation, repeated symbols, and domain context. Preserve original language, capitalization, punctuation, and line breaks when they carry meaning.

Never guess a material symbol. Flag ambiguous pairs such as 0/O, 1/l/I, i/j, hyphen/minus, Latin/Greek lookalikes, multiplication x/×, and superscript/subscript placement. Ask the user when context cannot resolve them reliably.

## Keep ordinary text native

Use editable PowerPoint text for Chinese, English, numerals, symbols, and mixed-language labels. Keep a sentence or label in one text container unless the source requires separately movable units. Apply bold, italic, underline, color, and other supported formatting to text runs instead of splitting objects.

Use the shape’s own text body for a labeled node. Use separate text boxes only when the text has independent geometry, layering, rotation, or edit behavior.

Do not convert text into images, SVG outlines, glyph paths, or per-character shapes.

## Author mathematical content by complexity

For simple expressions, use native text and supported character-level formatting. Keep variables italic, operators and function names such as max and min upright, and Greek letters as actual Unicode characters. Use real minus, multiplication, comparison, and infinity symbols where appropriate.

For superscripts, subscripts, stacked notation, or vertical text, verify current runtime support before relying on it. Follow the probe rules in [capability-matrix.md](capability-matrix.md); do not infer support from visual appearance in an intermediate render.

For fractions, radicals, integrals, sums, matrices, piecewise functions, multi-level scripts, hats, or other complex notation, prefer a genuine editable Office Math equation object only when the current toolchain can create and preserve it through PPTX export.

If true equation authoring is unavailable, pause and offer narrow alternatives:

1. Editable linear text with reduced typographic fidelity
2. Unicode super/subscript characters when the needed glyphs exist
3. A validated pre-existing native equation object supplied in a template
4. A non-editable formula image as a last resort

State which semantic, visual, and editability properties each option preserves. Obtain explicit approval before using an option that changes the reconstruction contract.

Use this priority for mathematics: semantic correctness, native editability, readable typesetting, source-like placement, then pixel-level similarity.

## Select fonts deliberately

Identify the source typeface only when visual evidence is credible. Check that the chosen font is installed in the build and rendering environments and contains every required glyph. Prefer the exact font when available; otherwise choose a stable, visually close font with suitable language coverage.

Prefer common PowerPoint-safe Chinese sans-serif fonts such as Microsoft YaHei or DengXian when they are available and visually appropriate. Prefer Cambria Math for mathematical runs when it improves glyph consistency. Do not force these defaults when the source clearly uses another style.

Use separate font runs only where script coverage or mathematical styling requires them. Preserve weight, italicization, size hierarchy, alignment, line spacing, and text direction without manufacturing unnecessary objects.

Treat font substitution as an explicit deviation when it materially changes geometry, identity, or line breaks. Explain the substitution and request approval before finalizing a fidelity-critical replacement.

## Validate typography

Render at full slide size and inspect every dense or mathematical region. Check for missing glyph boxes, fallback-font jumps, clipped accents, incorrect scripts, altered minus signs, broken line wrapping, overflow, and baseline drift.

Reopen or reimport the exported PPTX when possible and verify that text remains text and equations remain the promised object type. Do not accept a visually correct preview if export silently converted the content into paths or images.
