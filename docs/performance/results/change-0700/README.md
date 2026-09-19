# Change 0700 — use the bounded substring search for MCE detection

`performance_claim: none`. Disposition: **rejected**. The production codec is
restored to the 0700 baseline; the five focused regression tests are retained.
The candidate changes the MCE early predicate from a byte-by-byte `windows`
scan to the existing `memchr::memmem` search. It improves marker-free inputs,
but repeated real workflows are 1.5–3.4% slower and the tiny marked control
adds about 45% (about 210 ns) plus 216 bytes of stack reservation. Those costs
outweigh the marker-free gains, so the timings do not describe a retained
production gain.

The packet binds the complete source diff for
`crates/litchi-ooxml-common/src/mce/codec.rs` and
`crates/litchi-ooxml-common/src/mce/tests.rs`. `baseline.json` records the
parent revision, accepted constraints, build inputs, and the 602-file source
census. The focused tests are run against the unchanged baseline, the measured
candidate, and the restored codec plus retained tests. All three receipts bind
the exact 101 passing `mce::` tests to their source maps. The candidate witness,
source patch, rejection transition, and final tests-only source state remain
auditable.

## Reproduce

Use a disposable checkout with this packet at the same relative location.
Preserve the root `Cargo.lock`, copy `workspace-Cargo.lock` into the
disposable checkout root, and use the Rust toolchain recorded in
`environment.json`. CPU affinity and scratch paths are host-specific. Do not
refresh dependencies while reproducing the packet.

Before applying the candidate source, run the control check and baseline
focused tests, then build the native, allocation, refusal, and oracle probes:

```sh
python3 docs/performance/results/change-0700/prepare-control.py
python3 docs/performance/results/change-0700/check-control.py
python3 docs/performance/results/change-0700/run-focused.py baseline
python3 docs/performance/results/change-0700/build.py baseline
python3 docs/performance/results/change-0700/build-refusal.py baseline
python3 docs/performance/results/change-0700/build-oracle.py baseline
python3 docs/performance/results/change-0700/assembly.py baseline
python3 docs/performance/results/change-0700/processor-assembly.py baseline
python3 docs/performance/results/change-0700/measure.py baseline
python3 docs/performance/results/change-0700/measure-allocations.py baseline
python3 docs/performance/results/change-0700/measure-refusal.py baseline
python3 docs/performance/results/change-0700/profile.py baseline
```

Apply only the two paths recorded by `source-diff.patch`, format the
candidate, and freeze the exact source diff before measuring it:

```sh
git apply docs/performance/results/change-0700/source-diff.patch
cargo fmt --all
python3 docs/performance/results/change-0700/source-diff.py
python3 docs/performance/results/change-0700/run-focused.py candidate
```

The baseline focused receipt is created before applying either source path.
The candidate receipt is created after both paths are present and formatted.
This ordering prevents the baseline test result from silently exercising
candidate code. Both receipts must report the frozen 101 passing `mce::` tests.
The exact command sequence used for the batch is retained in the logs.

Build and measure the candidate with the same lockfile and inputs:

```sh
python3 docs/performance/results/change-0700/build.py candidate
python3 docs/performance/results/change-0700/build-refusal.py candidate
python3 docs/performance/results/change-0700/build-oracle.py candidate
python3 docs/performance/results/change-0700/assembly.py candidate
python3 docs/performance/results/change-0700/processor-assembly.py candidate
python3 docs/performance/results/change-0700/measure.py compare
python3 docs/performance/results/change-0700/measure-allocations.py candidate
python3 docs/performance/results/change-0700/measure-refusal.py compare
python3 docs/performance/results/change-0700/profile.py candidate
python3 docs/performance/results/change-0700/binary-sizes.py
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
python3 docs/performance/results/change-0700/oracle/corpus.py \
  /path/to/litchi-0700-bin/baseline-oracle \
  /path/to/litchi-0700-bin/candidate-oracle \
  /path/to/checkout/test-data \
  /path/to/packet/oracle-results
python3 docs/performance/results/change-0700/measure-oracle-controls.py
python3 docs/performance/results/change-0700/measure-oracle-real-controls.py
python3 docs/performance/results/change-0700/measure-oracle-declaration-controls.py
python3 docs/performance/results/change-0700/measure-marker-controls.py
```

The three secondary control families use independent 300-sample, ten-warmup
AA/ABBA measurements. Do not overlap Cargo builds, perf collection, or native
measurements. Generate deterministic summaries after all measurements:

```sh
python3 docs/performance/results/change-0700/summarize.py
python3 docs/performance/results/change-0700/summarize-allocations.py
python3 docs/performance/results/change-0700/summarize-refusal.py
python3 docs/performance/results/change-0700/report-metrics.py
```

The seven integration gates and six evidence gates are mandatory. Run them
after the measurement lane is terminal and before cleanup:

```sh
python3 docs/performance/results/change-0700/run-integration.py
python3 docs/performance/results/change-0700/quality-summary.py
python3 docs/performance/results/change-0700/run-evidence.py
```

Only after the complete seven-row integration result and six-row evidence
result are terminal, run the isolated follow-up. It writes only the separate
`followup/` tree and `followup-*` receipts, leaving the initial native,
refusal, allocation, profile, oracle, and marker-control results untouched:

```sh
python3 docs/performance/results/change-0700/measure-followup.py
python3 docs/performance/results/change-0700/audit_followup.py
```

The follow-up uses four ABBA process legs, 300 samples, and ten warmups. It
contains six native cases (`one-real`, `noop-real`, `two-real`, `one-control`,
`one-generated`, and `one-notes-poi`) plus the complete ten-case refusal
matrix. Its independent audit recomputes every native phase statistic,
refusal total statistic, pair delta, metadata identity, source/binary hash,
command, output path, and sample index before the full audit can pass.

After the follow-up audit passes, record the rejection transition. This keeps
the measured candidate codec witness and restores only the production codec;
the five added tests remain in the final checkout. Then run the retained
focused proof before the final audit and cleanup:

```sh
python3 docs/performance/results/change-0700/reject-candidate.py
python3 docs/performance/results/change-0700/run-focused.py retained
python3 docs/performance/results/change-0700/audit.py
python3 docs/performance/results/change-0700/cleanup.py --apply
python3 docs/performance/results/change-0700/seal.py
```

The audit recomputes raw statistics, checks source and binary hashes, verifies
the candidate `memmem` predicate and absence of the old `windows` predicate at
the MCE early return, replays the source patch, checks all three focused test
receipts, requires all 13 native cases, ten refusal cases, 78 allocation
comparisons, 1,920 oracle invocations, the isolated follow-up, and all seven
plus six gate rows. Gate receipts remain bound to the candidate source map
after the production codec is restored. It is run before cleanup and again
after the owned target, binary, profile, and marker-control paths are removed.
The root Cargo lockfile and unrelated scratch are preserved.

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
document. Exact oracle parity and focused tests are required before any timing
result is considered eligible for retention.
