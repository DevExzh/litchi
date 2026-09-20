# Change 0701 — lower-setup MCE namespace precheck

`performance_claim: none`. Disposition: **retained with scoped evidence**.
The candidate keeps the existing MCE dispatch semantics while moving the
namespace search into a private `#[inline(never)]` helper. The helper uses
`memchr` to find only valid first-byte positions, checks the complete URI with
`starts_with`, and advances one byte after a failed match. The process function
uses that helper for the exact early predicate; no `memmem` candidate is part
of this packet. The candidate is retained only for the scoped evidence-backed
review described below; this packet makes no general performance claim.

The packet binds the complete source diff for
`crates/litchi-ooxml-common/src/mce/codec.rs` and
`crates/litchi-ooxml-common/src/mce/tests.rs`. `baseline.json` records the
parent revision, accepted constraints, build inputs, and the 602-file source
census. Focused tests cover the original-codec baseline and the measured
candidate; the current focused census is 104 passing `mce::` tests on each
side. The candidate witness, source patch, helper assembly, and all source and
binary bindings remain auditable. If review rejects the candidate later,
the rejection transition is added as a separate, explicit packet step.

## Reproduce

Use a disposable checkout with this packet at the same relative location.
Preserve the root `Cargo.lock`, copy `workspace-Cargo.lock` into the
disposable checkout root, and use the Rust toolchain recorded in
`environment.json`. CPU affinity and scratch paths are host-specific. Do not
refresh dependencies while reproducing the packet.

Run the control check, then build every baseline artifact while the checkout
still has the original HEAD source and 101-test file. Do not apply either
source path before this baseline build and assembly lane:

```sh
python3 docs/performance/results/change-0701/prepare-control.py
python3 docs/performance/results/change-0701/check-control.py
python3 docs/performance/results/change-0701/build.py baseline
python3 docs/performance/results/change-0701/build-refusal.py baseline
python3 docs/performance/results/change-0701/build-oracle.py baseline
python3 docs/performance/results/change-0701/assembly.py baseline
python3 docs/performance/results/change-0701/processor-assembly.py baseline
python3 docs/performance/results/change-0701/measure.py baseline
python3 docs/performance/results/change-0701/measure-allocations.py baseline
python3 docs/performance/results/change-0701/measure-refusal.py baseline
python3 docs/performance/results/change-0701/profile.py baseline
```

Apply only the tests path from `source-diff.patch`, then run the baseline
focused census. This receipt intentionally runs the original codec with the
three added tests; the baseline build witnesses and their 101-test source map
remain unchanged:

```sh
git apply --include=crates/litchi-ooxml-common/src/mce/tests.rs \
  docs/performance/results/change-0701/source-diff.patch
cargo fmt --all
python3 docs/performance/results/change-0701/run-focused.py baseline
```

Apply the codec path, format the candidate, freeze the exact two-path source
diff, and run the candidate focused census:

```sh
git apply --include=crates/litchi-ooxml-common/src/mce/codec.rs \
  docs/performance/results/change-0701/source-diff.patch
cargo fmt --all
python3 docs/performance/results/change-0701/source-diff.py
python3 docs/performance/results/change-0701/run-focused.py candidate
```

Both focused receipts must report 104 passing `mce::` tests. The baseline
release/build witnesses still bind the original 101-test source map, while
the baseline and candidate focused receipts share the 104-test census. The
audit binds each distinction and retains the exact command sequence.

Build and measure the candidate with the same lockfile and inputs:

```sh
python3 docs/performance/results/change-0701/build.py candidate
python3 docs/performance/results/change-0701/build-refusal.py candidate
python3 docs/performance/results/change-0701/build-oracle.py candidate
python3 docs/performance/results/change-0701/assembly.py candidate
python3 docs/performance/results/change-0701/processor-assembly.py candidate
python3 docs/performance/results/change-0701/measure.py compare
python3 docs/performance/results/change-0701/measure-allocations.py candidate
python3 docs/performance/results/change-0701/measure-refusal.py compare
python3 docs/performance/results/change-0701/profile.py candidate
python3 docs/performance/results/change-0701/binary-sizes.py
```

The primary native matrix contains 13 workflows: one, noop, and two-target
edits over the real, mechanism-control, generated, and notes inputs. It uses
four baseline and two candidate process legs in AA/ABBA order, with 100
samples and five warmups per process. Refusal coverage contains ten cases with
the same schedule. Allocation diagnostics use separate instrumented binaries
and are interpreted independently of native timing. All raw TSV and stderr
files remain the evidence boundary.

Run the differential oracle before timing controls. It covers 192 deterministic
inputs, five capability profiles, and both frozen binaries (1,920 invocations)
with exact output, ownership, and report comparisons:

```sh
python3 docs/performance/results/change-0701/oracle/corpus.py \
  /path/to/litchi-0701-bin/baseline-oracle \
  /path/to/litchi-0701-bin/candidate-oracle \
  /path/to/checkout/test-data \
  /path/to/packet/oracle-results
python3 docs/performance/results/change-0701/measure-oracle-controls.py
python3 docs/performance/results/change-0701/measure-oracle-real-controls.py
python3 docs/performance/results/change-0701/measure-oracle-declaration-controls.py
python3 docs/performance/results/change-0701/measure-marker-controls.py
```

The three secondary control families use independent 300-sample, ten-warmup
AA/ABBA measurements. Do not overlap Cargo builds, perf collection, or native
measurements. Generate deterministic summaries after all measurements:

```sh
python3 docs/performance/results/change-0701/summarize.py
python3 docs/performance/results/change-0701/summarize-allocations.py
python3 docs/performance/results/change-0701/summarize-refusal.py
python3 docs/performance/results/change-0701/report-metrics.py
```

The seven integration gates and six evidence gates are mandatory. Run them
after the measurement lane is terminal and before cleanup:

```sh
python3 docs/performance/results/change-0701/run-integration.py
python3 docs/performance/results/change-0701/quality-summary.py
python3 docs/performance/results/change-0701/run-evidence.py
```

Only after the complete seven-row integration result and six-row evidence
result are terminal, run the isolated follow-up. It writes only the separate
`followup/` tree and `followup-*` receipts, leaving the initial native,
refusal, allocation, profile, oracle, and marker-control results untouched:

```sh
python3 docs/performance/results/change-0701/measure-followup.py
python3 docs/performance/results/change-0701/audit_followup.py
```

The follow-up uses four ABBA process legs, 300 samples, and ten warmups. It
contains six native cases (`one-real`, `noop-real`, `two-real`, `one-control`,
`one-generated`, and `one-notes-poi`) plus the complete ten-case refusal
matrix. Its independent audit recomputes every native phase statistic,
refusal total statistic, pair delta, metadata identity, source/binary hash,
command, output path, and sample index before the full audit can pass.

After the follow-up audit passes, record the scoped retention decision and run
the candidate audit before cleanup:

```sh
python3 docs/performance/results/change-0701/audit.py
python3 docs/performance/results/change-0701/cleanup.py --apply
python3 docs/performance/results/change-0701/seal.py
```

The audit recomputes raw statistics, checks source and binary hashes, verifies
the candidate helper shape and assembly, replays the source patch, checks the
focused test receipts, requires all 13 native cases, ten refusal cases, 78
allocation comparisons, 1,920 oracle invocations, the isolated follow-up, and
all seven plus six gate rows. It is run before cleanup and again after the
owned target, binary, profile, and marker-control paths are removed. The root
Cargo lockfile and unrelated scratch are preserved.

## Interpretation

Native timers cover capture, working clone, text editing, commit, and apply on
fresh prepared packages. Loading, initial opening, target search, save,
serialization, and semantic oracle work are outside that timer. Profiles have
their own setup and assertion denominator. Allocation requested-byte totals
include replacement sizes for reallocations and must not be added to realloc
bytes a second time.

The host is shared and warm; the packet makes no cold-cache, quiescence, RSS
bound, native Office, or general throughput claim. Every >5% trigger remains
visible in the raw comparison receipts. The marker-control archive changes
the namespace URI and is a mechanism control, not a semantically equivalent
document. Exact oracle parity and focused tests are required before a timing
result is eligible for scoped retention.

## Retained evidence

The initial and follow-up audits passed. Both focused receipts contain 104
passing `mce::` tests, all 1,920 oracle invocations have exact output,
ownership, and report parity, and all seven integration plus six evidence gates
passed.

Follow-up mechanism-control deltas range from −11% to −15%. The real cases
show +0.07%/+0.68% for `one-real` and −0.61%/+0.34% for `two-real`. The first
baseline `noop-real` comparison shows −6.87% drift, so it does not support a
7% speedup claim. Initial real-case results show a 1.1%–2.8% cost, with the
observed tails retained in the evidence. The tiny marked marker-control case
shows a small gain and remains scoped to that control.
