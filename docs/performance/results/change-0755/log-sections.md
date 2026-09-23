# Log paragraphs for change 0755

Three blocks, one per shared log. Each is written to be prepended as the newest
section of its file.

## For `HOTSPOTS.md`

## 0755 — a nested empty text run panicked `set_shape_text`

[0755](0755-pptx-nested-text-run-panic.md) is a correctness fix found by change 0743's review, not a performance change. An empty `a:t` inside an open `a:t` passed the scene reader, which refused only a nested start tag, and the opened text-run locator recorded it as a span starting inside its parent's span; `set_shape_text` then sliced the slide with a reversed range and, with `panic = "abort"`, killed the process. The scene reader and the locator now refuse any nested `a:t` with the typed error semantic text already used, the rewrite checks its spans before emitting, and every slice the opened edit path takes from span data goes through a checked helper. It costs nothing measurable (instructions within ±0.05% on every semantic PPTX region). Open follow-up: quick-xml does not count a skipped UTF-8 byte-order mark in `buffer_position`, so every span of a BOM-prefixed part is three bytes early and its text edits are refused (never corrupted); the fix belongs in the crate's position helpers. [Evidence](results/change-0755/README.md).

## For `REPORT.md`

## 0755 — PPTX text edits refuse nested text runs instead of panicking

[0755](0755-pptx-nested-text-run-panic.md) is retained with `performance_claim: none`. The reported slide (`<a:t>M<a:t/>y`) now yields a typed `Error::Invalid` from `set_shape_text`, `set_shape_texts` and `Slide::shapes`, as it already did from `Slide::text`; a 4,230-variant mutation test of text-body markup drives every text verb under `catch_unwind` with no panic, and fails at variant 1,050 without the fix. No repository deck has a nested `a:t` in a slide-like part, so no fixture's result changes; per-operation instructions move by at most 0.05%. [Evidence](results/change-0755/README.md).

## For `GOAL_AUDIT.md`

## 0755 — malformed input is refused, never a panic, and readers of one element agree

[0755](0755-pptx-nested-text-run-panic.md) closes a gap between three readers of the same DrawingML text element: semantic text refused a nested `a:t` in both forms, the scene reader only in one, and the opened edit path turned the other into an out-of-order span and a slicing panic, which `panic = "abort"` makes a process abort. The fix aligns the refusals and, beyond the reported site, routes every span slice of the opened edit path through a checked helper, so the path is panic-free by construction rather than by the ordering arguments each call site relied on. Capture still accepts such a deck, because it validates the package graph and notes roots rather than text runs; every verb that reads the runs refuses it. [Evidence](results/change-0755/README.md).
