# 0823 — direct `Scene::Scanner::scan` trial packet

This packet records a completed, rejected before/after trial for the single
candidate in `crates/litchi-pptx/src/shape/reader.rs`. The candidate replaces
the `NsReader` transport in `Scene::Scanner::scan` with `Reader` plus the
existing `NamespaceResolver`; the review must preserve namespace scope
transitions, pending-pop timing, event/error mapping, limits, and refusal
ordering. All 342 admitted reports/7,106 samples replay successfully. The best
eligible paired p50 improvement is 2.394%, below the frozen 3% benefit gate;
allocations are unchanged. Production is restored to the exact baseline.
See [the report](../../0823-pptx-scene-reader-workflow.md), `disposition.json`,
and `results-review.md` for the decision and limitations.

The synthetic probe is copied from the sealed 0813 source. Its timed corpus
has six shapes (`tiny`, `medium`, `large`, `vendor`, `unicode-vendor`, and
`valid-4attr`) crossed with `capture`, `commit`, and `lifecycle`. Its
allocator wrapper carries the current process-wide live-byte conservation test
fix; the timed corpus and output oracle are unchanged. The real probe is the
packet-local 0822 real-file edit probe with its schema/base identity refreshed
for this trial. Both manifests are independent Cargo packages and use the
checked-in root dependency path.

The static packet binds the full production and `tools/perf-baseline` source
censuses at freeze time, the root/rustfmt/tool lock copies, all 35 architecture
inputs, host/toolchain metadata, the two probe trees, the candidate before/
after/patch archive, and the three unrelated workspace-file hashes. The
owned target is `/home/zhuhe/code/litchi-target-0823`. The baseline quality
driver may leave only its `quality` child in that target before `freeze.py` and
the first release build.

The baseline quality receipt omits `.cargo/config.toml` from its source-state
map because its command enumerator did not include that path. Freeze therefore
checks every covered production path against the receipt and separately binds
`.cargo/config.toml` byte-for-byte to the declared base commit. This is a
metadata-only custody amendment; the config defines an unused lint alias and
does not alter the trial commands.

The original root-owned execution sequence is below. Reproduction requires a
fresh checkout of the pinned base, this packet's frozen input files, and fresh
output directories; the capture drivers deliberately refuse existing outputs
and a different HEAD. The archived failed native lane must never be merged
into the admitted matrix.

```text
python3 -B docs/performance/results/change-0823/quality.py before
python3 -B docs/performance/results/change-0823/quality.py probes
python3 -B docs/performance/results/change-0823/freeze.py
python3 -B docs/performance/results/change-0823/build.py before
python3 -B docs/performance/results/change-0823/capture.py qualification before
# Root reviews the before qualification and then applies the candidate.
python3 -B docs/performance/results/change-0823/quality.py after
python3 -B docs/performance/results/change-0823/build.py after
python3 -B docs/performance/results/change-0823/capture.py qualification after
python3 -B docs/performance/results/change-0823/capture.py native
python3 -B docs/performance/results/change-0823/capture.py allocation
python3 -B docs/performance/results/change-0823/codegen.py
```

`build.py` performs four serial release builds per leg (synthetic and real,
native and allocator) with offline locked Cargo, two build jobs, opt-level 3,
debug level 1, thin LTO, one codegen unit, unwind panic, and incremental
compilation disabled. Every child workload is pinned to CPU 12 and wrapped by
`/usr/bin/time` for an RSS receipt. Every failed child writes its receipt and
log before the driver raises; an existing output is never overwritten or
retried.

Qualification runs one allocation-instrumented sample for each of the 19 cases
on each leg (38 reports total). It is an observer lane: each report must match
the retained synthetic fixture/output/semantic oracle or the pinned real-file
identity, and every verification field is checked together with the bytes,
hashes, semantic digests, and allocation accounting. Native timing uses six
counterbalanced `before/after` blocks, 30 samples and three warmups, for 228
reports and 6,840 samples. Allocation uses two counterbalanced blocks, three
samples and no warmup, for 76 reports and 228 samples.

The benefit gate is unchanged in form: at least one eligible public workflow
must have median after/before at most 0.97 with bootstrap upper endpoint below
1.00, using 10,000 resamples, seed 823823, endpoints 250 and 9749. Eligible
rows are the twelve synthetic commit/lifecycle rows and real/direct; synthetic
capture rows remain negative controls. Any of all 19 rows with bootstrap lower
endpoint above 1.05 vetoes adoption. Allocation medians may not increase calls,
allocated bytes, net live bytes, or peak above entry; an RSS increase above 5%
requires review. No universal, cross-format, cold-cache, tail-latency, RSS, or
historical speedup claim follows from this packet.

The first native lane stopped after an unexpected agent commit changed HEAD;
`native-interrupted-0` retains its original logs, nine report artifacts, and
commit patch. Root preserved all files, restored the prior local HEAD, and
restarted the entire lane without changing frozen inputs or analyzing timings.
Three pre-freeze probe-quality failures and one offline reader-schema failure
retain logs and source snapshots. See `execution-notes.md` for the full record.

The owned target was removed after all eight binaries were hash-verified:
9,687 files and 3,238,577,403 logical bytes. The cleanup receipt substitutes
exact binary witnesses during offline replay. No native executable is needed
for these checks:

```text
python3 -B docs/performance/results/change-0823/analysis.py --check
python3 -B docs/performance/results/change-0823/raw_audit.py --check
python3 -B docs/performance/results/change-0823/profile_basis.py --check
python3 -B docs/performance/results/change-0823/validate.py --final
python3 -B docs/performance/results/change-0823/seal.py --check-head
```

The seal binds every packet file, the main report, and the five performance
indexes to the exact staged or committed change set from the measured base.
The unrelated format review, unified API design, and empty matrix file remain
untouched and excluded. No iWork implementation is changed.
