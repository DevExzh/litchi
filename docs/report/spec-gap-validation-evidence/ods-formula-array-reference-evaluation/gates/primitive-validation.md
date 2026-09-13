# Value evaluation primitives

The private scalar bridge and geometry helper have completed isolated
validation. The scalar bridge passed six tests and warning-denied Clippy;
geometry passed three standalone tests. These results do not establish that
the value evaluator, worksheet adapter, or their integration tests pass.

The scalar bridge reuses the existing scalar kernels and preserves Empty
until the required operand type is known. Its tests cover owned-text
reservation release after a failed argument, subsequent kernel reuse,
constants, pending frames, mismatched function nodes, cancellation, eager
argument exclusion for lazy/sequence functions, and the explicit Empty
operand profile. The independent assessment is in `scalar-bridge-review.md`.

Geometry covers checked nonempty cuboids, arithmetic overflow, half-open
intersection, bounding cuboids, and projection with unequal extents and
different origins. It does not decide formula implicit-intersection policy.

## Reproduction

The scalar tests used a clean disk worktree of baseline commit
`67360b209c5b428161b27f1f08650169d4997739`. The exact workspace dependency
lock is retained as `primitive-workspace.Cargo.lock`; place a copy at the
worktree's `Cargo.lock`. The baseline evaluator file hash is recorded in both
scalar receipts. Append the module declaration from
`scalar-bridge-test-context.txt` to that evaluator file, adjusting its absolute
path to the reviewed `value/scalar.rs` source if the checkout location differs.
The declaration is under `cfg(test)` and exposes the private bridge to the
baseline kernels without importing the unfinished value VM.

Run the commands from `scalar-bridge-test.json` and
`scalar-bridge-clippy.json` using their recorded environment. Both commands
must finish before restoring the baseline evaluator file. The recorded runs
restored its exact bytes in a `finally` block, and checked the bridge source
hash before and after each command. Build targets and temporary files were
kept on disk outside `/tmp` and `/var/tmp`.

For geometry, use a temporary Rust wrapper containing a single private module
declaration whose `#[path = "..."]` points to the reviewed `value/geometry.rs`.
Compile the wrapper with `rustc --edition 2024 --test`, then run the resulting
binary. `geometry-test.compiler`, `geometry-test.source-sha256`,
`geometry-test.status` and `geometry-test.stdout` retain the observed compiler,
source hash and results. The temporary wrapper and executable were removed.

The complete candidate must still pass the normal package gates without any
wrapper. Candidate performance comparisons and resolver-backed end-to-end
tests remain required before accepting the array/reference implementation.
