# Coalesced PPTX command transactions

Consecutive distinct direct ordinal shape text/geometry sets now share one
native primary Edit and one commit in atomic `File::batch` and `File::edit`.
This implements command coalescing, rather than treating one selector over many
shapes as many commands. Source and eager owners use static dispatch.

The coalescer shares property admission with ordinary PPTX setters. A run keeps
one detached native slide scene and bounded ordinal indexes; target lookup does
not rescan the shape list for every command. The source owner remains immutable.
Plans retain only native positions and coerced assignments. Native staging uses
bounded contiguous same-slide/family calls, splitting by both the native
256-replacement capacity and its aggregate text-byte limit. All calls publish
through one commit. No native capacity or validation policy was widened.

Repeated targets close the run, and partial geometry is resolved again against
the resulting revision. Queries, identity/nested paths, reads, metadata,
structural operations and other families use existing barriers/handlers.
Best-effort mode isolates every item. Native text and bounds equality determine
per-item changed reporting. Failed staging, publication or admission leaves
the caller's source untouched.

Wire-order replay locates deferred staging failures. Commit-only failures use
prefix isolation. Earlier pending failures win over a following malformed path,
property or command payload. PPTX native proofs are boxed in transaction history,
like DOCX/XLSX proofs, to avoid inflating every cross-format history entry.

## Verification

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,107 tests/doctests.
The no-feature suite passed 419. Ten focused coalescing regressions also passed
after the final proof-storage adjustment. They cover 1,000 separate commands
with one native commit, text/mixed properties, exact no-ops and per-item change
flags, eager/source replay and inverse, source immutability, repeated-target
partial geometry, read/query/metadata/structural barriers, semantic limits,
best-effort isolation, late malformed transforms and error-index precedence.
A generated pair of texts whose sum exceeds the native 8 MiB per-call cap
verifies byte-based chunking while preserving one unified commit.

Combined facade all-target Clippy and standalone harness Clippy passed with
warnings denied. Native leaf code was unchanged. The rewritten base remains
an ancestor of this branch; no rebase is required.

## Measurement

The new `unified-pptx-command-batch` harness compares one direct native Edit
using bounded calls against an atomic batch of distinct direct-path commands.
Both owner modes and text/text+x workloads are measured. Inputs, command vectors
and native replacements are prepared outside timing; opening, warm-up,
serialization, output equivalence checks and teardown are excluded. Every
unified sample must report one native commit and one result per input command.

The run uses CPU 0, counts 1/10/100/1,000/10,000, and three ABBA rounds (six
samples per leg). No build or test runs concurrently. This is scoped evidence;
allocation/RSS, real-world decks, cold workloads, other command families,
remaining formats and complete goal acceptance are still required.

The first matrix attempt stopped at eager 10,000 commands before emitting its
final JSON. The harness now retains completed results and structured failures,
continues other cases, and exits nonzero when any case fails. The complete
matrix was rerun; failed scenarios are not counted as successful measurements.

The eager 10,000-shape refusal occurs during native Edit capture, before any
setter chunk. `SlidePart::finish_from_processed()` requests notes-root
classification, whose scanner applies `notes::MAX_NODES = 100_000` to slides.
The generated 10,000 text boxes exceed that ceiling. The conformance retry
reduces the underlying `Error::Limit` to `None`, and the root proof reports an
invalid root/namespace instead. This is an open native resource-policy/error
classification finding, not malformed benchmark XML. A follow-up should use
resource-specific slide-root node admission aligned with the existing bounded
slide readers, retain notes' own ceiling and the other byte/depth/attribute
bounds, and preserve limit evidence through classification. No limit was
relaxed in this change.

[Machine-readable evidence](0156-pptx-command-batch-results.json) contains all
samples, corpus hashes, build/source-state provenance, binary size/SHA-256 and
the two structured eager failures. The matrix exits 1 because those cases fail.

| Owner | Properties | Commands | Native median ms | Unified median ms | Ratio |
|---|---|---:|---:|---:|---:|
| eager | text | 1 | 0.305 | 0.463 | 1.516 |
| eager | text_x | 1 | 0.334 | 0.498 | 1.492 |
| source | text | 1 | 0.053 | 0.055 | 1.033 |
| source | text_x | 1 | 0.081 | 0.082 | 1.014 |
| eager | text | 10 | 0.414 | 0.646 | 1.561 |
| eager | text_x | 10 | 0.587 | 0.824 | 1.404 |
| source | text | 10 | 0.269 | 0.272 | 1.009 |
| source | text_x | 10 | 0.435 | 0.441 | 1.014 |
| eager | text | 100 | 1.460 | 2.324 | 1.592 |
| eager | text_x | 100 | 2.988 | 3.847 | 1.287 |
| source | text | 100 | 2.344 | 2.407 | 1.027 |
| source | text_x | 100 | 3.820 | 3.922 | 1.027 |
| eager | text | 1000 | 31.511 | 38.575 | 1.224 |
| eager | text_x | 1000 | 82.211 | 88.919 | 1.082 |
| source | text | 1000 | 77.476 | 77.847 | 1.005 |
| source | text_x | 1000 | 128.052 | 129.828 | 1.014 |
| source | text | 10000 | 7289.616 | 7303.048 | 1.002 |
| source | text_x | 10000 | 12141.836 | 12286.120 | 1.012 |

The 1,000-command unified/native ratios range from 1.005 to 1.224, meeting the
2x comparison target for these workloads in this run. This is not complete
batch acceptance: source 10,000-command latencies grow about 94x over 1,000
for 10x the commands. Both native and unified legs exhibit this behavior,
consistent with full-slide reparsing for each bounded native call. The leaf
needs staging that reuses its parsed slide and performs bounded publication
without repeated whole-slide work. The eager 10,000-node-policy failure must
also be fixed. One-command `File::batch` has extra coordinator admission cost;
these figures do not replace record 0155's ordinary `File::set` measurements.

```sh
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0156-target cargo test --offline -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0156-target cargo clippy --offline -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption --all-targets -- -D warnings
CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0156-target cargo build --offline --release --manifest-path tools/perf-baseline/unified-pptx-command-batch/Cargo.toml
taskset -c 0 /tmp/litchi-unified-0156-target/release/unified-pptx-command-batch-baseline --counts 1,10,100,1000,10000 --rounds 3
```

Set build-provenance environment values as in the harness README when
regenerating the evidence. The full goal remains active.

## Repository and cleanup

Scoped formatting and diff checks passed. The artifact audit found no prohibited
evidence or compiled binary among 690 changed/untracked paths. Both staged-file
guards printed nothing. Scoped staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on its
read-only filesystem; no commit was created.

All verification and measurement processes ended before cleanup removed the
turn-owned build output, temporary logs/raw JSON and generated root lockfile.
The small evidence JSON retains all successful samples and failed scenarios;
no compiled binary is retained in the repository. No temporary worktree was
created. Inherited files and the base checkout were preserved.
