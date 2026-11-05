# One PPTX slide publication across bounded chunks

Records 0156–0157 measured nonlinear homogeneous command batching because each
bounded native setter call reparsed and republished the entire slide. The native
primary eager and source edits now accept `set_shape_text_chunks` and
`set_shape_property_chunks`; focused source slide edits expose the same
operations without a slide selector. The facade submits each contiguous
same-slide shape run through these operations, including alternating text-only
and geometry updates.

Each input chunk retains the existing 256-target and aggregate text-byte
limits. Chunk count and total selection count are bounded by the default shape
reader's 100,000-shape ceiling; total replacement text is also bounded by the
output-byte budget. Admission precedes descriptor flattening. All targets are
resolved against one immutable slide state, and duplicate or overlapping
selections across chunks are refused. Text/geometry validation, lexical
preservation and semantic readback use the shared native codecs. One candidate
is published after the complete operation succeeds. Existing single-chunk
setters and their limits remain available.

The facade preserves command-level results, no-op reporting, barriers and
wire-order failure attribution. New benchmark routes ending in
`chunk_publication_v2` compare the bulk native operation against one atomic
facade batch containing a separate command per shape. The prior
`command_batch_v1` harness remains unchanged.

## Verification

The native PPTX/encryption suite passed 1,688 tests/doctests. The combined
DOCX/XLSX/PPTX/encryption facade suite passed 1,108. New native coverage checks
1,000-shape byte equivalence with sequential bounded setters, 10,000-shape
edits, source immutability, eager/source/focused exact no-ops, cross-chunk
duplicates and invalid payloads, empty chunk admission, per-chunk and aggregate
bounds, managed reservation refusal/retention, and transported replay/inverse.
Two roughly 4 MiB text values test aggregate text beyond one chunk's 8 MiB
budget, for both text and combined property operations. The facade regression
checks all 1,000 alternating text/geometry results in one parsed output scene.
Existing wire-order error and transaction-barrier tests remain green.

Native/facade and standalone harness all-target Clippy passed with warnings
denied. The full PPTX glossary profile passed 70 vocabulary bans. Repository
boundaries passed for 66 workspace packages and 268 dependency declarations.

## Measurement

The new harness ran all 20 cases on CPU 0, counts 1/10/100/1,000/10,000,
both owner modes, text/text+x, and three ABBA rounds (six samples per leg).
No build or test ran concurrently. All warm-up outputs matched exact bytes;
every timed unified result retained one native commit and one changed result
per input command. Each native leg calls the chunked setter once, passing up
to 40 chunks for 10,000 commands, and commits once.

The first matrix run rejected the large cases because a copied harness guard
still expected one native call per chunk. That guard now requires one call for
the complete chunked operation in warm-up and timed samples; harness Clippy
and release build passed again before the complete matrix was rerun.

| Owner | Operation | Commands | Native median (ms) | Unified median (ms) | Unified/native |
|---|---|---:|---:|---:|---:|
| eager | text | 1 | 0.304 | 0.468 | 1.539 |
| eager | text_x | 1 | 0.335 | 0.500 | 1.492 |
| source | text | 1 | 0.054 | 0.055 | 1.029 |
| source | text_x | 1 | 0.081 | 0.083 | 1.017 |
| eager | text | 10 | 0.410 | 0.650 | 1.585 |
| eager | text_x | 10 | 0.593 | 0.837 | 1.410 |
| source | text | 10 | 0.271 | 0.274 | 1.009 |
| source | text_x | 10 | 0.444 | 0.454 | 1.022 |
| eager | text | 100 | 1.474 | 2.329 | 1.580 |
| eager | text_x | 100 | 3.042 | 3.909 | 1.285 |
| source | text | 100 | 2.363 | 2.416 | 1.022 |
| source | text_x | 100 | 3.880 | 3.935 | 1.014 |
| eager | text | 1000 | 11.714 | 18.752 | 1.601 |
| eager | text_x | 1000 | 26.385 | 33.382 | 1.265 |
| source | text | 1000 | 23.277 | 23.690 | 1.018 |
| source | text_x | 1000 | 38.095 | 38.578 | 1.013 |
| eager | text | 10000 | 112.640 | 181.115 | 1.608 |
| eager | text_x | 10000 | 263.025 | 333.861 | 1.269 |
| source | text | 10000 | 234.160 | 237.837 | 1.016 |
| source | text_x | 10000 | 384.054 | 387.548 | 1.009 |

| Owner | Operation | Native 10k/1k | Unified 10k/1k |
|---|---|---:|---:|
| eager | text | 9.616 | 9.658 |
| eager | text_x | 9.969 | 10.001 |
| source | text | 10.060 | 10.040 |
| source | text_x | 10.081 | 10.046 |

The 1,000-command unified/native ratios are 1.013–1.601; 10,000-command
ratios are 1.009–1.608. All meet the scoped 2x comparison target. Unified
10k/1k growth is 9.66–10.05x for 10x commands, consistent with linear growth
in this generated corpus. This replaces record 0156's roughly 94x source
growth finding for these command workloads. New route names distinguish the
native bulk operation from the prior repeated-call comparator. One-command
`File::batch` overhead is not ordinary `File::set` overhead; record 0155's
single-set evidence remains a separate scenario.

[Machine-readable evidence](0158-pptx-chunk-publication-results.json) retains
all raw samples, corpus identities, build command, source revision,
dirty-source manifest digest and binary size/SHA-256. The compiled binary
is never committed. Larger/wider corpora, allocation/RSS, other command
families, selector-fanout consumers still using sequential bounded setters,
remaining formats and complete goal acceptance remain open. Managed source
admission is conservative; these timings do not establish managed 10k capacity.

## Reproduction

```sh
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0158-target cargo test --offline -p litchi-pptx --features encryption
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0158-target cargo test --offline -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR=/tmp/litchi-unified-0158-target CARGO_NET_OFFLINE=true cargo build --offline --release --manifest-path tools/perf-baseline/unified-pptx-chunk-publication/Cargo.toml
taskset -c 0 /tmp/litchi-unified-0158-target/release/unified-pptx-chunk-publication-baseline --counts 1,10,100,1000,10000 --rounds 3
```

Set provenance environment variables as described in the harness README.
Native corpus tests use the three LibreOffice fixture copies listed in record
0157; the original base checkout remains untouched.

## Repository and cleanup

The rewritten base remains an ancestor; no rebase was required. Scoped
formatting and diff checks passed, and both required staged-file guards printed
nothing. An audit of 705 newly committed and pending paths found no prohibited
large evidence or compiled binaries.

Commits made outside this agent's tool calls advanced the branch during the
turn. Native implementation is included in `724a37462f`, facade implementation
in `723f8c13eb`, harnesses in `4aced1adff`, and this measurement record in
`c63a0ad4fb`. The sole source difference after the measured source revision is
`e8e91b4ec4`, which reordered XLSX module declarations. Substituting that file's
measured-revision content exactly reproduces the recorded source manifest;
PPTX and harness code are unchanged. The evidence preserves this provenance.

This agent's own scoped staging still failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on its
read-only filesystem. The implementation and initial measurement record had
already been committed by the concurrent process; this agent cannot create
commits under the current filesystem permissions.

All verification and measurement jobs ended before cleanup. Hash-checked
turn-owned fixture copies, the generated root lockfile, temporary logs/raw JSON
and build output were removed. The reduced evidence retains every timing
sample and binary provenance. No temporary worktree was created; inherited
files and the base checkout were preserved. The full goal remains active.

