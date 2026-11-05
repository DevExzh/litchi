# 0362ul89k — PPTX F5 dead-code repair receipt

Date: 2026-10-25. This is a source-only follow-up to the actual F5
all-features/all-targets failure at
`237d6801a641b85b4477c432a26ac53e531987ed`. It records the assigned
dead-code repair; it does not claim that the retained F5 build succeeded.

## Retained failure and immutable repair

The retained actual capture is outside Git at
`(deleted-coord-artifact)`.
The all-features/all-targets diagnostic log is 49,586 bytes with SHA-256
`e3004cf8f79ec32a2dd2b082f4204dc1efe5a673dab7df97dcb63f9203aa06ea`.
Its `diagnostics-index.json` is 16,760 bytes with SHA-256
`775ff0196bcbcf8ec70f1bb66ee969419181664f262944dd3775f308c09f235e` and
contains 113 `dead_code` diagnostics in total. The assigned subtree contains
49 of those diagnostics; other PPTX folders remain owned by their respective
workers.

The repair is immutable at commit
`c551c4e266a90b670c17774b594428b5d643d4b1`, based directly on the retained
237 source. Its external patch is
`(deleted-coord-artifact)/pptx-f5-dead-code-237-c551c4e266.patch`,
22,621 bytes, SHA-256
`75bd206bb6689d9ef16232cd7333e48e23154692706166f510b7d1cd800c14f3`.
The patch contains only the 13 assigned source paths. `rustfmt --check`,
`git diff --check`, assigned-scope checks, the no-`3rdparty/` audit, and the
repository size guards passed. No Cargo, rustdoc, build, or native command was
run for this follow-up.

## Replacement and capability audit

The patch removes only private default twins with no production consumers:
change-tracking codec/package defaults, collaboration graph validation,
notes/media codec defaults, comments legacy commit helpers, and the related
notes defaults. It marks wrappers that are genuinely used only by tests with
`cfg(test)` (animation host scanning, snapshot construction, comments,
media-part, and notes test helpers). No `allow(dead_code)`, public alias, or
default-limit fallback was added.

Every production selected-policy route remains available and its consumer
continues to use the corresponding `*_with_limits` path: change-tracking
state/shape mutation and package application; collaboration presence,
threading, and graph validation; presentation comment author/comment edits;
bounded media-part codec/package operations; notes XML/resource scanning,
package loading/removal/clearing, graph validation and proof-aware loading;
and related-notes validation/eager processing. The retained test wrappers are
not substitutes for those production routes.
