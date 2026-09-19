# Zen desktop visual system

Zen is a local **Operate** surface built with WinUI 3: the contemporary native Windows 11 language, not an imitation in HTML or a legacy Windows 10 control set.

## Direction contract

- Use the Windows App SDK visual system: Mica backdrop, system title bar, `NavigationView`, `InfoBar`, native `TextBox`, `NumberBox`, `ProgressRing` and `Button` controls.
- Keep the task funnel short: choose an operation, select a source, optionally choose the output, run and read an explicit result.
- Stretch the task surface across the available window with 24–32 epx gutters. On wide desktop windows, use a larger file workspace on the left and a narrower summary/action pane on the right.
- Treat source and destination as one file workflow, separated by a native divider. Keep status and the primary action together in the summary pane so the whole window has a clear functional balance.
- Group operations in the left navigation using the user’s vocabulary: Convertir, Extraire and Image.
- Use WinUI system typography, theme resources and state animations. No custom web typography, CSS cards, web navigation chrome, dark marketing sidebar or decorative gradients.
- Reveal parameters only for crop and rounded-corner operations, so ordinary conversions remain focused.
- Use Mica as the window surface and WinUI’s card resources only to create functional separation, not a dashboard appearance.

## Accessibility and states

- Preserve native keyboard navigation, focus visuals, high-contrast compatibility and system scaling.
- Surface completion, errors and warnings through `InfoBar` text; color is never the sole signal.
- Disable the primary action and show a native `ProgressRing` while a local operation runs.
- Keep all UI labels and result messages in French.
