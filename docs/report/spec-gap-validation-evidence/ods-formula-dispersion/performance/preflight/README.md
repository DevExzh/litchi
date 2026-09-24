# Untimed preflight receipt

The profile harness was built in release mode with the retained gate lockfile
and run with `--preflight-only`. This mode parses and evaluates each selected
fixture through the same correctness, resolver, budget, and formula-error
paths used before measurement, then exits without calling the timed evaluator
or `/usr/bin/time`.

The candidate isolated worktree passed all 103 named cases in the evaluate
phase. The detached baseline at
`55e147bfa0676ce6ecdc609efc682b98568b8a5f` passed all 17 matched controls.
The candidate list refusals for VAR, VARP, and STDEVP completed without a
resolver read; the harness source and verifier retain those zero-read bounds
for the later capture.

The release build targets and copied harnesses were removed after this
receipt. These are correctness/read-bound checks only; they are not timing
samples and are excluded from performance results.
