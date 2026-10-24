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

The rerun passes at `0f007c4d74b59fbc8a38f3181236d947403c3e7b` with the
repaired monitor `f9cf4aa6d5`: strict all-target/all-feature Clippy over
`litchi-cli`, `litchi-ops`, `litchi` and `litchi-py` exits zero in 27.094 seconds.
The source stays clean, three post-exit samples find no owned processes or
builders, and the target remains 4,706,526,684 bytes. The `-q` success log is
empty; its exit and provenance are recorded in the source-bound state receipt.

[The reduced receipt](results/0362ugka/consumer-check.json) binds the exact
command, locks, policy, external plan/state/log identities and post-exit
observations. This is one selected check, not the full 53-step gate, and carries
no native qualification claim.
