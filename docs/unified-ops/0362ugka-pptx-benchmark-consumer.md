# 0362ugka — PPTX find/replace benchmark consumer

The combined consumer check at `0d35a83a31` found one missed migration in
`litchi-ops/examples/find_replace_bench.rs`: the PPTX edit's `slides()` getter
now returns a fallible projection. Iterating the `Result` produced a reference
to the complete slide slice instead of individual slides and failed to compile.
The benchmark now propagates the getter error before iteration. It preserves
the selected slide parts, edit operations and measured boundary.

The diagnostic log is outside the repository:
`/home/zhuhe/code/litchi-worktrees/.coord/budget-validation/composed-lazy-consumers-0d35a83a31/01-clippy-all-features-attempt-1792811253378323215-2180562.log`
(505 bytes, SHA-256
`dbf46af8711e49cfaab1f1a196435149256b99739975b031c21383964bc43371`).
The gate monitor separately encountered a vanished-process `ESRCH` race and
recorded the attempt as blocked; the actual child log retains this compiler
failure. Both issues require repair before the combined check can pass.

This correction carries no native qualification claim. The combined consumer
check will be rerun on the corrected source with the repaired monitor.
