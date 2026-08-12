# Change 0081: RTF batch and sequential-text evidence

Date: 2026-08-12

Measurement base reported by the raw runs:
`21a214b68928f3a0d819a52c46dcfd1d98a4eaa5`

Status: accepted harness evidence; no production-code change

## Decision and scope

Two opt-in RTF scenarios are now independently selectable:

- `rtf_semantic_one_percent_edit_save` changes `ceil(1%)` of the ordinary
  body paragraphs in one generated large plain-RTF document, commits, and
  serializes the published snapshot; and
- `rtf_semantic_text_to_sink` converts the already-open semantic body to UTF-8
  through `Document::write_text_to` and a bounded forward-only non-seek sink.

The one-percent case deliberately accepts only plain RTF. CP-1252 and LZFu
remain read-only for this edit because the retained transaction refuses an
unsupported changed transport rather than normalizing it. The text-sink case
accepts plain, CP-1252, and LZFu because all three can be parsed into the same
read-only semantic facade. The producer-watermark corpus is excluded because
it is not an eligible ordinary body-text workload.

This tranche does not change `litchi-rtf` or any other production crate. It
selects the already-landed `Edit::replace_body_paragraph_texts` API for the
one-percent harness case after measuring it against repeated calls to the
existing scalar `replace_paragraph_text` API. The paired release binaries use
identical production code; their only intentional difference is the retained
small [scalar comparator patch](../results/rtf-batch-sequential-0081-scalar-comparator.patch).
No iWork/IWA code or evidence is involved.

## Machine, toolchain, and source state

- Ubuntu 24.04.4 LTS, Linux 6.8.0-101-generic, x86-64 little-endian.
- AMD EPYC 9575F 64-Core Processor; the environment exposed 12 logical CPUs.
- Every retained measurement was pinned with `taskset -c 2`; the raw process
  reports therefore see one available logical CPU.
- `rustc 1.95.0 (59807616e 2026-04-14)`, LLVM 22.1.2, and
  `cargo 1.95.0 (f2d3ce0bd 2026-03-21)`.
- Cargo `release` profile and Rust system allocator.
- `RUSTFLAGS`, `CARGO_BUILD_TARGET`, and `LD_PRELOAD` were unset.
- `/proc/sys/kernel/perf_event_paranoid` was `1`.

The raw reports record revision `21a214b6` and `git_worktree_dirty: true`.
Both alternatives were frozen seconds apart from the same shared worktree,
with no production-source change between builds. The dirty state is disclosed
rather than represented as a clean-revision result. The exact intentional
comparator change is stored, and the batch source is the retained harness
state in this change.

Measured frozen-binary identities:

| Harness strategy | SHA-256 |
|---|---|
| one atomic batch | `4fa0ce9dc067af58d50ee8fdbe167c32f1691027ff847f51065e3e0f6995d59a` |
| repeated scalar edits | `d772d928f75db5e230967f78b9d5d3250d35b7d68b2e132aea2b069db9844e78` |

## Corpus and correctness oracle

The deterministic `litchi-rtf-semantic-v2` large plain corpus contains 10,000
ordinary body paragraphs, 499,999 visible UTF-8 bytes, and 540,051 source
bytes. Its SHA-256 is
`957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e`.

The shared `semantic_update_indices` rule selects 100 strictly increasing,
evenly spaced source-relative positions. Both strategies request identical
replacement strings. The deterministic changed RTF contains 540,151 bytes
and has SHA-256
`d040328cb691fc5ec65192477688f4a9a4275a8b62fa354a2fdb68d739786d8f`.

After every timed iteration, the harness checks the complete output bytes,
reopens the result, verifies all 10,000 paragraph texts and the flattened body,
applies the exact-source patch, and applies its inverse to restore the original
source bytes. A focused late-out-of-range batch regression proves staging
atomicity and source-snapshot sharing after failure. Stable large-output hashes
are also pinned by a focused harness regression.

## Measurement build and comparator recipe

The measured batch binary was built with the following commands. Its SHA-256
identifies the retained measurement artifact; it is not an expected output of
a later rebuild because the raw run records a dirty worktree without a complete
tree patch or tree hash.

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
cargo build --locked --release \
  --manifest-path tools/perf-baseline/Cargo.toml
cp tools/perf-baseline/target/release/litchi-perf-baseline \
  /tmp/litchi-perf-rtf-batch-0082b
sha256sum /tmp/litchi-perf-rtf-batch-0082b
```

Starting from the same batch harness state, the scalar comparator was built
with the stored source delta, then the batch harness was restored:

```sh
git apply --check \
  docs/performance/results/rtf-batch-sequential-0081-scalar-comparator.patch
git apply \
  docs/performance/results/rtf-batch-sequential-0081-scalar-comparator.patch
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml
cargo build --locked --release \
  --manifest-path tools/perf-baseline/Cargo.toml
cp tools/perf-baseline/target/release/litchi-perf-baseline \
  /tmp/litchi-perf-rtf-scalar-0082b
sha256sum /tmp/litchi-perf-rtf-scalar-0082b
git apply -R \
  docs/performance/results/rtf-batch-sequential-0081-scalar-comparator.patch
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml
```

The patch changes only the measured staging call from one bounded batch to a
loop of scalar calls. Expected-output construction, timing boundaries,
commit/save work, and verification are identical.

## One-percent ABBA protocol

The retained CPU-2 order was batch A, scalar A, scalar B, batch B. Every leg
used 30 warmups followed by 200 retained samples, for 400 samples per state.
The exact command shape was:

```sh
taskset -c 2 BINARY \
  --case rtf_semantic_one_percent_edit_save \
  --semantic-shape large \
  --rtf-variant plain \
  --samples 200 \
  --warmup 30 \
  --json OUTPUT.json
```

The raw reports are [batch A](../results/rtf-batch-sequential-0081-abba-batch-a.json),
[scalar A](../results/rtf-batch-sequential-0081-abba-scalar-a.json),
[scalar B](../results/rtf-batch-sequential-0081-abba-scalar-b.json), and
[batch B](../results/rtf-batch-sequential-0081-abba-batch-b.json). The
[machine-readable summary](../results/rtf-batch-sequential-0081-summary.json)
contains pooled statistics, commands, artifact digests, output hashes, and
timing boundaries.

Source parsing, update selection, replacement construction, expected-output
construction, and sink reservation happen before `Instant::now`. The interval
contains edit construction, staging, commit with complete candidate reopen,
publication, one shared snapshot-handle clone, and RTF serialization into the
pre-reserved non-seek sink. Exact-byte comparison, semantic reopen/readback,
patch replay, inverse restoration, and sink reporting happen after elapsed
time is captured.

## One-percent results

Per-leg results:

| Leg | p50 | p95 | Mean | Mean 95% confidence interval |
|---|---:|---:|---:|---:|
| batch A | 5.470 ms | 6.096 ms | 5.520 ms | 5.473-5.568 ms |
| scalar A | 7.089 ms | 7.632 ms | 7.159 ms | 7.120-7.199 ms |
| scalar B | 7.147 ms | 8.065 ms | 7.295 ms | 7.227-7.362 ms |
| batch B | 5.371 ms | 5.781 ms | 5.416 ms | 5.390-5.443 ms |

Pooled results:

| Metric | Repeated scalar | Atomic batch | Batch delta |
|---|---:|---:|---:|
| samples | 400 | 400 | — |
| p50 | 7.118 ms | 5.413 ms | **-23.95%** |
| p95 | 7.885 ms | 5.918 ms | **-24.94%** |
| mean | 7.227 ms | 5.468 ms | **-24.33%** |
| mean 95% confidence interval | 7.187-7.267 ms | 5.441-5.496 ms | disjoint |

Both ordered comparisons favor the batch strategy materially. Between-leg p50
and mean drift are below 2% for each state. The evidence therefore clears the
five-percent materiality gate for using the existing batch API in this named
generated workload. It is not evidence of a newly implemented production
optimization. Although the raw harness format includes its standard p99
field, this record makes no p99 claim.

## Sequential semantic text baseline

The separate text case used the retained batch binary, ten warmups, and 50
samples per large variant:

```sh
taskset -c 2 /tmp/litchi-perf-rtf-batch-0082b \
  --case rtf_semantic_text_to_sink \
  --semantic-shape large \
  --rtf-variant plain,byte1252,lzfu \
  --samples 50 \
  --warmup 10 \
  --json /tmp/rtf-text-to-sink-large-0082b.json
```

Only `Document::write_text_to` is timed. Document parsing, expected UTF-8
construction, exact sink reservation, complete byte/report comparison, and
full semantic verification are outside timing. No RTF package serialization
or LZFu decoding is inside the interval.

| Variant | Source SHA-256 | UTF-8 output | Output SHA-256 | p50 | p95 | Mean (95% CI) |
|---|---|---:|---|---:|---:|---:|
| plain | `957645f9…c6e02e` | 499,999 B | `f122900c…736abdc` | 0.745 ms | 0.843 ms | 0.757 ms (0.748-0.767) |
| CP-1252 | `7157437b…50e` | 529,999 B | `b6fb32c1…a5872` | 0.747 ms | 0.811 ms | 0.752 ms (0.745-0.759) |
| LZFu | `e293574f…5176` | 499,999 B | `f122900c…736abdc` | 0.750 ms | 0.784 ms | 0.754 ms (0.748-0.760) |

The complete source and output digests are in the summary. Plain and LZFu
emit the same semantic UTF-8. The sink observes 19,999 bounded writes; the
largest write is 49 bytes for plain/LZFu and 52 bytes for CP-1252. These are
baseline observations only, with no before/after optimization claim.

## Correctness and evidence gates

The retained harness source passes:

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
cargo check --locked --manifest-path tools/perf-baseline/Cargo.toml
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  semantic_rtf -- --nocapture
cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets -- -D warnings
git apply --check \
  docs/performance/results/rtf-batch-sequential-0081-scalar-comparator.patch
```

The focused smoke covers deterministic all-variant capability selection,
plain/CP-1252/LZFu sequential conversion, watermark refusal, the one-percent
batch's atomic late failure, complete reopen, patch replay, inverse restoration,
and stable large evidence hashes. The repository copies of all five raw JSON
reports parse successfully; canonical sorted JSON comparisons match their
source reports. Their repository SHA-256 values are pinned in the summary.

## Limitations

- The latency result covers one deterministic generated 10,000-paragraph
  plain-RTF corpus. It does not generalize to formatted, media-bearing,
  real-producer, malformed, protected, or opaque-rich documents.
- CP-1252 and LZFu are conversion baselines only; this tranche does not enable
  editing those transports.
- The text conversion numbers begin after document parsing and transport
  decode, and do not measure RTF serialization.
- No allocation, peak-RSS, hardware-counter, cold-cache, concurrent, or
  remote-source conclusion is made.
- The acceptance worktree was dirty. The paired comparison remains controlled
  because production source was unchanged between frozen builds and the exact
  comparator delta is retained, but a later clean rebuild can have a different
  binary SHA-256 as unrelated dependency source evolves.
- Raw reports contain the harness-standard p99 statistic. It is deliberately
  excluded from the accepted result because this protocol makes no p99 claim.
