# Batch 0970 retained DOC A/A source-variability review

Status: fresh correction to v1, 2026-10-08. v1 is preserved unchanged at
`source-variability-review-v1.md`. This review launched no Cargo command,
build, benchmark, workload, or reducer and did not change the checkout or any
XLS draft. It does not reopen the 0969 readiness decision and makes no current
revision, candidate, speedup, slowdown, or host-cause claim.

## Scope and evidence custody

The retained 0969 A/A wave is the only runtime evidence considered here. Its
frozen plan is `/home/zhuhe/litchi-performance-0969/plans/plan-0969-aa-v6.json`
(SHA-256 `e853030180d7464c06cf1ed5d525957ce2d50514b8337a1b168392374d00b5df`).
It used the ordinary baseline executable from 0967:

```text
/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile
SHA-256 0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb
```

The exact binary/build/source identity is bound by the retained 0967
receipts, rather than by a new build for this diagnostic:

* build summary `/home/zhuhe/litchi-performance-0967/receipts/build-0967-cfb-v1.json`, SHA-256 `d1de24c3a310d7f9d6e4c0819a7aac9287138715d835b73d1f062739d37dd5d6`;
* build/ELF configuration audit `/home/zhuhe/litchi-performance-0967/receipts/build-config-elf-audit-0967-v1.json`, SHA-256 `a6b30eed1ea78dffd376b23cf7cea7db82bf53301f4ce9c0f3077da15d6853f3`;
* source baseline manifest used by that build
  `/home/zhuhe/litchi-performance-0967/receipts/source-baseline-manifest-v1.json`, SHA-256 `bf247c041f74b257dda4333f26d9f64b8dda25f639729a692bef39156f3a756e`.

That immutable source export has revision
`893072657ad8b1d726bfce351b5ec2e98f084e94`, 10,694 files, aggregate file
hash `75e84ea561e11dc0070610ca54396a6f8a77696222cbb6248b22fb2e91273610`,
and records `build_started: false` and `workload_started: false` in the
source receipt. The build summary records the ordinary binary as 47,909,224
bytes with the SHA above and the locked offline `+1.95.0`, `-j2` build
configuration.

The bounded 0969 reuse review is also retained at
`/home/zhuhe/litchi-performance-0969/review/baseline-reuse-v1.json`, SHA-256
`dbd362dae8e31dfe5ff54d6af5c37883df829bede3a5e5fa19e04337e0a8312e`. Its
selected runtime/harness/configuration/fixture scope matches the historical
baseline and explicitly forbids timing or allocation-equivalence inference.
The 0969 per-file manifest
`/home/zhuhe/litchi-performance-0969/manifests/source-file-manifest-v1.json`
has SHA-256
`99e65d43e0feed89133eec60b095130d68058a071a02e8a15b05dea01209bb6c`, 924
selected entries, and separate `baseline_sha256` values. A future runner must
verify those baseline entries against the historical `source-baseline` root;
it must never admit the candidate entries. These receipts bind the historical
binary and source files only. They make no claim about the current checkout or
any 0969 candidate.

The fixtures are the exact byte-bound controls:

| fixture | bytes | SHA-256 | descriptive CFB shape |
| --- | ---: | --- | --- |
| `FloatingPictures.doc` | 335,360 | `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d` | 15 directory entries, including nested `ObjectPool` storage and its streams |
| `NoHeadFoot.doc` | 26,112 | `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48` | 6 directory entries, no `ObjectPool` |

The terminal capture receipt
`/home/zhuhe/litchi-performance-0969/capture/aa-v4/capture-stage-v4.json`
has SHA-256
`59485d2c231f2c56e39300acfc5916c5a14ceea91918ebe0c2706fc7fe666bc3` and
records `readiness_fail`, `no_retry`, `real_lanes_started: false`, and
`formal_candidate_launch_permitted: false`. The governing A/A receipt
`/home/zhuhe/litchi-performance-0969/capture/aa-v4/aa-readiness-v4.json` has
SHA-256
`6bb110145d0583be3ff12c3c6e1ada6611be6b5396f55f24698bb0358bb51efe` and
uses 12 pairs per cell, seed `969004`, 10,000 bootstrap resamples, and upper
limit `1.05`.

## What the raw packets establish

The matrix contains 24 cells: two fixtures, edits `1/2/8/32`, and retention
`zero/default/release`. Each cell has 12 fresh AB/BA pairs, six in each order;
there are 576 child packets and 46,080 retained measured samples. Every packet
has the same ordinary binary SHA and source-manifest SHA for its cell, a zero
return code, `timed_out: false`, no launch error, zero-byte stderr, no
`validation_error`, and 80 finite elapsed samples. The report's correctness
object is identical across the packet set and all correctness values are true.

For each cell, the output tuple `(output_sha256, output_bytes,
output_capacity)` is constant across all 24 child reports. The output is also
exactly equal to the prepared reference output in every report. The four
output tuples per fixture correspond to the four edit counts; retention does
not introduce an output mismatch. All ordinary reports set `timing_claim` to
true and `allocation` to null. The runner's outer wall duration and RSS files
are retained, but they are outside the in-process elapsed metric and are not
allocation counters.

The readiness result is `readiness_fail`: all 12 `docfloat` cells fail their
cell-level A/A decision, all 12 `docnohf` cells pass, and 42 of the 96
order/statistic blocks are over the frozen bound. This is a complete
same-binary A/A failure under the declared protocol. It is not evidence that a
candidate is slower, and it does not identify the cause of the baseline
variation.

The process-level spread is materially different between the two fixtures.
The following ranges are the minimum and maximum child-report percentile
among the 24 reports in each cell; each child percentile is computed from its
80 measured samples:

| cell | child p50 range | child p95 range |
| --- | ---: | ---: |
| `docfloat/e1/default` | 0.972–1.244 ms | 0.985–1.274 ms |
| `docfloat/e32/default` | 7.373–8.130 ms | 7.958–8.632 ms |
| `docnohf/e1/default` | 0.114–0.119 ms | 0.125–0.132 ms |
| `docnohf/e32/default` | 0.999–1.030 ms | 1.014–1.047 ms |

The order effect is not one fixed direction. For example, in
`docfloat/e1/default`, AB first and second child p50 means are about 1.094 ms
and 1.135 ms, while BA first and second are about 1.112 ms and 1.031 ms. In
`docnohf/e1/default`, the corresponding means are about 0.116 ms, 0.116 ms,
0.116 ms, and 0.117 ms. This weakens a simple universal first/second or
AB/BA explanation, but it does not distinguish allocator, scheduler, cache,
or frequency effects.

Within-child spread is also larger for `FloatingPictures`: for the two
representative default cells, the mean sample standard deviation is about
41.4 µs at edit 1 and 203.9 µs at edit 32, versus about 5.4 µs and 9.6 µs for
`NoHeadFoot`. Thus the retained evidence shows both a broad fresh-child
spread and broader 80-sample tails for the complex fixture. It does not prove
that the fresh-process state is the source of either effect.

There is no support in these packets for a JSON serialization mismatch as the
explanation. Serialization, report hashing, and post-sample validation occur
after the native clock; the reports have the expected sample count, stable
output identity, and identical correctness results. This rules out the
specific explanation that JSON encoding changed the measured elapsed values.
It does not rule out a bug in an analysis implementation, so any future
analysis must continue to validate packet schema and sample custody before
computing ratios.

## Source path and grounded hypotheses

The source anchors below are from the retained historical baseline export.
Their hashes are included to bind the code that produced the executable; they
are not a current-checkout or candidate-source claim. Relevant file hashes are:

| file | SHA-256 |
| --- | --- |
| `tools/perf-cfb-profile/src/lib.rs` | `b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc` |
| `tools/perf-cfb-profile/README.md` | `ca2206b276a0c84a840945c61853108a23addfe376aebe5c06d27695e1f20ad5` |
| `crates/litchi-doc/src/body_text.rs` | `0e8cb6351be3327322882ae4035230cd241aed0b6208cb8589e7d76692d3d786` |
| `crates/litchi-doc/src/tracked_revision/package.rs` | `a04da643b093ec9724c9ac57ed8fbfdf2c2fedfc16eade105c91c32df7c9a446` |
| `crates/litchi-ole-common/src/object/editor.rs` | `58913ec7125acb061aa5547a2f9a674cd314a1eb8bf3e90bba2df62392bf106c` |
| `crates/litchi-ole-common/src/object/codec.rs` | `25e4035e739788e29f10ed4e4ba77d4187221c26f789783d06e473ada6add146` |
| `crates/litchi-cfb/src/writer/core.rs` | `a5ead3dded13baa8bdaa350742bfe79185e39cc2f16f86e433701e939fb2b19c` |

The timing boundary is explicit in
`tools/perf-cfb-profile/src/lib.rs:290-330,332-358,388-410,686-765,937-1018`:

* `run_case` clones the source into an owned `Vec`, calls
  `Snapshot::open_bounded`, creates one public `Edit`, applies all repeated
  paragraph replacements and any retention release, commits, copies the
  committed snapshot bytes into another `Vec`, and drops the local values.
* `timed_case` starts `Instant` immediately before that lifecycle and stops
  immediately after it. The preparation routes, public-reader checks, output
  hashes, per-sample validation, and JSON serialization are outside the clock.
* The allocator branch uses the same operation wrapper under
  `alloc_metrics::region`, but deliberately emits no timing claim. The
  ordinary 0969 reports have no allocation counters.

The README states the same boundary at `tools/perf-cfb-profile/README.md:41-47`.
Therefore the measured number includes the lifecycle's source and package
work, but not report serialization or the outer child-launch wall time.

The DOC owner performs repeated complete construction in the timed path. At
`crates/litchi-doc/src/body_text.rs:652-698`, `Snapshot::open_bounded` owns the
source, opens the strict `RevisionEditor`, validates the complete CFB through
the package reader, and retains the source bytes. At
`body_text.rs:1326-1334`, `Snapshot::editor` clones the source bytes and opens
another revision editor for the staged transaction. At
`body_text.rs:1447-1472`, a commit either takes a retained render or clones the
editor and renders it. At `body_text.rs:2086-2108`, a changed candidate is
opened again through `Snapshot::open_bounded_with_policy` before the final
snapshot and patch are returned.

The lower tracked-revision path makes the retention difference concrete.
`crates/litchi-doc/src/tracked_revision/package.rs:3020-3056` clears the old
render, rebuilds the changed `WordDocument`/table/data stream replacements,
asks the package editor for a validated rendered result, and retains that
`Vec` only when its capacity is within the configured ceiling.
`package.rs:2654-2659` moves that retained render at finalization or calls the
package renderer when no render is held. Consequently, zero retention, default
retention, and release-after-stage retention can perform different allocation
and rendering work while still producing the same final bytes. The measured
harness intentionally includes those route differences.

The CFB publication path is also allocation and layout heavy. In
`crates/litchi-ole-common/src/object/editor.rs:934-1009`, a changed package
uses copy-through when admissible or constructs a fresh layout render; the
candidate editor then reopens and captures the emitted CFB. In the inspected
codec source, `crates/litchi-ole-common/src/object/codec.rs:27-64` binds a
rendered output to its package model, and `codec.rs:187-294` either captures
all stream payloads or, for the newer private receipt path, walks the final
CFB under the same limits and projects exact source stream allocations. That
receipt/projection code is part of the inspected historical baseline source
context; it is not an explanation for the retained 0967 executable's A/A
numbers, and no candidate path was measured.

`crates/litchi-cfb/src/writer/core.rs:78,320-341` shows the writer's
`HashSet`/`HashMap` storage and metadata structures. Rust process-seeded hash
state can change probe paths and allocation/layout cost in a fresh child. The
writer deliberately removes hash iteration from serialized storage ordering:
`core.rs:525-559` builds ordered effective storages and class IDs, and
`core.rs:1743-1777` sorts the explicit storage and CLSID paths before assigning
SIDs. `core.rs:2294-2308` documents that determinism requirement. This makes a
hash-seed-induced byte-order mismatch unlikely and leaves only a plausible
cost/allocation hypothesis, which the retained packets do not measure. The
writer's layout planner and fallback are at `core.rs:1631-1659`; a plan may
validate and emit through the adopted layout, or the writer may fall back to
from-scratch serialization.

The strongest source-grounded reading is therefore: `FloatingPictures` has a
larger nested CFB and the timed DOC lifecycle repeats more complete parsing,
rendering, retention, and final reopen work, so it has more opportunities for
allocator, hash-table, layout, cache, and scheduler sensitivity. The data
show that sensitivity, but they do not select randomized hash state,
allocation/layout pressure, CPU scheduling, frequency, or cache state as the
cause. No production change follows from this review.

## Chosen 0970 discriminating diagnostic, for a later root GO only

The selected experiment is a historical-baseline affinity diagnostic. It
reuses the exact ordinary executable that produced the 0969 A/A packets and
changes one launcher factor: the child CPU affinity. Rebuilding would change
another variable, so no build is part of this design. There is no allocator
lane, current-revision binary, candidate binary, source instrumentation, or
formal qualification rerun.

The four fixed contrast cells are `docfloat/e1/default`,
`docfloat/e32/default`, `docnohf/e1/default`, and `docnohf/e32/default`. Each
cell uses the existing harness command with `--warmup 5 --samples 80`, 12
pairs, six AB and six BA pairs, and serial fresh children. Each pair has two
children, one in each affinity arm, for 24 children per cell and 96 children
in total. Odd pairs run unpinned then pinned; even pairs run pinned then
unpinned. The exact binary, fixture hashes, source-baseline admission, and
report schema remain bound to the receipts above. No child is retried,
replaced, filtered, or selected after observing its measurements.

The unpinned arm inherits the full allowed CPU mask from a fresh host audit.
The pinned arm uses one CPU, derived as the greatest CPU in that same audited
mask. The current audit suggests CPU 31, but a later fresh audit is
authoritative and capture stops if the mask has changed or the selected CPU is
not allowed. The affinity capability and telemetry references already reviewed
for this design are:

* `/home/zhuhe/litchi-performance-0970/diagnostic-design-v3.md`, SHA-256 `95aa5817f95a038c32dfbd4af7906eb474df3546aa321f0adf72c1572e3b462b`;
* `/home/zhuhe/litchi-performance-0970/host-audit-v1.md`, SHA-256 `37e204363bee72eb7d3080300dd32ce5ce21d53be86dee2c9e726b086f6901c5`;
* `/home/zhuhe/litchi-performance-0970/telemetry-probe-v1.json`, SHA-256 `8cb97d377c6162c57e1ae092c7c8e33f9b2b23c38aa2d44f188d3831e20614dc`;
* `/home/zhuhe/litchi-performance-0970/affinity-probe-v1.json`, SHA-256 `0f2e97080d2ba9898a3c0c74660f3280a4cbb66e4546997fa93749d476ce6a17`.

Before any root GO, the runner must repeat process quiescence and host-mask
checks. Every packet must bind its PID, requested and observed affinity,
exact `execve` argv, cgroup path, binary/fixture/source-baseline hashes, and
unique pair/order/cell identity. `/usr/bin/time` and, if the frozen capability
is still valid, `/usr/bin/perf` sidecars may record user/system time, context
switches, migrations, faults, task-clock, cycles, and instructions. Their
scope must be recorded as outer-launch telemetry; the harness `Instant` region
stays unchanged.

The existing 0969 report validator remains authoritative for every child:
zero return code, no timeout or launch error, the exact report schema, 80
finite elapsed samples, all correctness/oracle fields true, stable
output/reference identity, and the expected semantic edit count. Sidecar
identity and finite-value checks are equally strict. Any missing, malformed,
non-finite, wrong-identity, or failed packet makes the diagnostic incomplete;
there is no partial result or retry. The historical `readiness_fail` and its
`1.05` bound remain unchanged. This affinity diagnostic has no adoption
threshold and cannot convert that result into a pass or authorize a candidate.

The analysis compares pinned and unpinned child p50/p95 distributions, the
within-child `p95/p50` tail, order-stratified values, and paired scheduler
telemetry. A complete report may say that the observed data are consistent
with an affinity-associated change or that no separating signal was observed.
Neither conclusion identifies a root cause. The following interpretations are
predeclared:

| observation | limited interpretation |
| --- | --- |
| Pinning reduces `docfloat` within-child tails or fresh-child dispersion and the telemetry moves in the same direction | supports a scheduler/CPU-placement contribution; it does not prove one |
| Pinning changes neither tails nor dispersion while output and correctness remain stable | leaves process hash/allocation/layout, cache, frequency, and timer effects unresolved |
| Telemetry changes but harness p50/p95 does not | records a host difference without a source-performance conclusion |
| Output/reference hashes, correctness, sample counts, or route identities differ | marks the diagnostic incomplete or a harness/source fault; it is not evidence for an optimization |

An allocator lane, phase observer, or deterministic-hash comparison is a
separate possible future experiment. It would require its own design and
custody review; none is part of this chosen 96-child diagnostic. In
particular, this experiment deliberately preserves the historical source and
binary so that affinity is the only planned factor.

## Disposition

The measured route has real, removable-looking work only after the operation's
existing validation and rendering semantics: repeated source/editor opens,
per-stage rendered publication, retention transitions, final candidate reopen,
and output copying. The retained evidence does not justify removing any of
that work or attributing it to randomized hash state. The correct 0970
recommendation is to keep the candidate held, preserve the 0969 readiness
failure, and, if root authorizes it, run only the exact historical-baseline
affinity diagnostic above before considering a separate source or allocation
experiment. No current-revision or candidate performance conclusion follows.
