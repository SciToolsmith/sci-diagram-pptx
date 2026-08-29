# Cross-Platform Compatibility

Read this file only for custom/freeform labels, fragile typography, complex formulas, a reproduced platform problem, or an explicit cross-platform request. The default goal remains one portable editable PPTX, not pixel identity across applications.

## Remove the risky structure first

- Put labels for custom geometry and freeforms in overlaid standard text boxes unless the geometry exports a usable text rectangle and a focused render verifies it.
- Resolve text layout before export with explicit fonts, sizes, margins, line breaks, and practical headroom. Roughly 8%–12% spare width and height can guide compact labels, but judge the actual render rather than enforcing a threshold.
- Do not rely on open-time AutoFit to choose a font size, resize a shape, or rewrap text.
- Use structured baseline runs for simple scripts and proven Office Math for complex notation. Do not treat platform variants as a way to avoid resolving ambiguous scientific content.

## Keep the validation claim bounded

Resolve portability risk through stable construction and the normal render and structure checks. Treat a render as evidence for the selected renderer only, not proof of pixel identity across applications.

Do not create platform-specific derivatives by default. Deliver one portable file and disclose any known compatibility limitation without turning application-specific testing into a reconstruction requirement.
