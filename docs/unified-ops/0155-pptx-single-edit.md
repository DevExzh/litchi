# One native eager edit per PPTX setter

Record 0154 found 45–51% eager overhead on one-shape selector writes. Inspection
identified a discarded native Edit in `Pptx::editable()`, followed by another
Edit for the actual setter. Each construction captures and validates the native
graph and creates transaction state.

The owner acquisition is now shared through a private `editable_owner()`.
The setter uses it and immediately constructs its actual native Edit, including
for empty properties and semantic no-ops. Admission-only callers retain
`editable()`, which still runs the native checks. Signature, retention, source
freshness and execution checks remain in owner acquisition. No native API,
validation policy or public facade shape changed.

A new regression uses a readable first slide followed by malformed XML. The
native eager Edit refuses that graph; unified empty and equal-text setters must
also refuse it and preserve exact input bytes. This protects against skipping
native admission when the request appears to do no work.

## Paired measurement

[Machine-readable evidence](0155-pptx-single-edit-results.json) records all ABBA
samples, corpus hashes, build command, source-state digest and binary SHA-256.
The same harness/routes as 0154 ran on CPU 0 with nine ABBA rounds (18 samples
per leg), release debug disabled and no concurrent build/test job. Preparation,
opening, serialization, output-equivalence checks and teardown are outside the
timer. Native/unified warm-up outputs passed equivalence checks in every case.

| Owner | Properties | Shapes | Native median ms | Unified median ms | Ratio |
|---|---|---:|---:|---:|---:|
| eager | text | 1 | 0.307 | 0.319 | 1.039 |
| eager | text_x | 1 | 0.335 | 0.352 | 1.051 |
| source | text | 1 | 0.053 | 0.057 | 1.075 |
| source | text_x | 1 | 0.080 | 0.084 | 1.052 |
| eager | text | 10 | 0.417 | 0.454 | 1.089 |
| eager | text_x | 10 | 0.588 | 0.628 | 1.069 |
| source | text | 10 | 0.265 | 0.277 | 1.046 |
| source | text_x | 10 | 0.434 | 0.443 | 1.022 |
| eager | text | 100 | 1.487 | 1.770 | 1.190 |
| eager | text_x | 100 | 2.985 | 3.283 | 1.100 |
| source | text | 100 | 2.356 | 2.385 | 1.012 |
| source | text_x | 100 | 3.848 | 3.909 | 1.016 |
| eager | text | 1000 | 31.625 | 34.238 | 1.083 |
| eager | text_x | 1000 | 83.852 | 87.175 | 1.040 |
| source | text | 1000 | 78.830 | 79.178 | 1.004 |
| source | text_x | 1000 | 130.261 | 131.496 | 1.009 |

One-shape ratios are 1.039/1.051 for eager text/mixed and 1.075/1.052 for source.
They meet the 10% small-write budget for these generated workloads in this run.
This does not establish complete performance acceptance. The 100-shape eager
text ratio is 1.190. Eager selector resolution still parses the selected scene;
its share of the remaining overhead has not been isolated. Counts above 256 still stage multiple native calls
and reparse the slide. Separate command coalescing, the 10,000-shape matrix,
linear scaling, cold/real-world workloads and allocation/RSS remain open.

Reproduction (populate provenance environment values as in the harness README):

```sh
CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0155-target cargo build --offline --release --manifest-path tools/perf-baseline/unified-pptx-shape-fanout/Cargo.toml
taskset -c 0 /tmp/litchi-unified-0155-target/release/unified-pptx-shape-fanout-baseline --counts 1,10,100,1000 --rounds 9
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0155-target cargo test --offline -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
```

The rewritten base remains an ancestor of this branch; no rebase is required.
The full goal remains active.

## Verification and repository state

The combined DOCX/XLSX/PPTX/encryption facade suite passed 1,097 tests/doctests,
including the new malformed-graph empty/equal setter regression. Scoped
formatting passed. The artifact audit found no prohibited evidence or compiled
binary among 682 changed/untracked paths. Both staged-file guards printed
nothing. Scoped staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on its
read-only filesystem; no commit was created.

PPTX facade all-target Clippy passed with warnings denied. Diff checks passed.
All processes ended before cleanup removed this turn's temporary build output,
raw measurement/log files and generated root lockfile. The small evidence JSON
retains all samples and binary provenance. No temporary worktree was created;
inherited files and the base checkout remain untouched.
