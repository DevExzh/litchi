# 0362ul89g — PPTX 61-scope compiler return

Date: 2026-10-25. This is a source-only compiler repair; no runtime,
performance, or qualification result is claimed.

The repair is based on frozen source `00a42a7d4c9bfead038e2e40c57a0c14291e5ecb`
and addresses only the assigned `litchi-pptx` 61-scope diagnostics. The
129,953-byte diagnostic log is retained outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-00a42a7d4c/01-pptx-all-features-check-attempt-actual-f5-01.log`
with SHA-256
`d7f08d8f7a99af9566b704d86cc9ca0066f28e745d2c0f5ac0cdd3e6ec11628f`.

The changes make the laser usage arithmetic and tag source intersection
constant-safe; restore crate-internal selected-span and classification loader
seams; migrate tag, classification, and Designer XML replacements to their
selected output limits; remove stale imports/reexports; and qualify the
`OnceLock` use. Test-only tag package consumers were migrated to the
crate-internal raw helpers. No default limit fallback was introduced.

Static checks completed on the fresh sparse worktree: `git diff --check`,
assigned callsite/reexport scans, no conflict markers, and no `3rdparty/` or
binary references in changed files. Cargo, rustdoc, and runtime tests were not
run as required by the coordinator.
