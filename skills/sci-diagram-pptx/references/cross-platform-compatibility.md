# Cross-Platform Compatibility

Read this file only for custom/freeform labels, fragile typography, complex formulas, a reproduced platform problem, or an explicit cross-platform request. The default goal remains one portable editable PPTX, not pixel identity across applications.

## Remove the risky structure first

- Put labels for custom geometry and freeforms in overlaid standard text boxes unless the geometry exports a usable text rectangle and a focused render verifies it.
- Resolve text layout before export with explicit fonts, sizes, margins, line breaks, and practical headroom. Roughly 8%–12% spare width and height can guide compact labels, but judge the actual render rather than enforcing a threshold.
- Do not rely on open-time AutoFit to choose a font size, resize a shape, or rewrap text.
- Use structured baseline runs for simple scripts and proven Office Math for complex notation. Do not treat platform variants as a way to avoid resolving ambiguous scientific content.

## Validate only the claimed environment

Use one real PowerPoint open or export only when a risk trigger applies and PowerPoint is accessible. Check for a repair prompt, missing objects, material reflow or clipping, displaced custom/freeform labels, and broken formulas. Record the actual environment; never imply that another platform was tested.

Create platform-specific derivatives only when the blocker is reproduced in real PowerPoint on each affected environment, portable construction cannot remove it, and every derivative is generated from the same canonical object data and validated in its named environment. Otherwise deliver one portable file and disclose the validation scope.
