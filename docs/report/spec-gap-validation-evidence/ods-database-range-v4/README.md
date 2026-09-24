ODS database-range metadata
===========================

The 22 source paths match independently reviewed V4. This batch adds bounded typed database-range read and source-backed editing, including source/filter/sort/subtotal metadata and qualified same-sheet range addresses. It preserves unknown or structurally unsupported markup on exact no-ops and refuses changed rewrites that would lose it. Database execution and refresh remain inert.

Root validation on the committed ODG/shared-ODF base passed 523 tests across 37 targets including doctests, 0 ignored, strict all-feature/all-target Clippy, warning-denied rustdoc and scoped formatting. Independent review verified parent grammar, nested filter groups, no-op preservation and atomic changed-edit refusals. Source and explicit build-lock hashes remained unchanged across root gates.

The clean integration uses git archive of 52519031b plus the exact feature paths. There is no overlap with primary pending work; all 17 committed ODG/shared-ODF feature paths are retained. The root Cargo.lock is an explicit untracked build input, not claimed as committed source.
