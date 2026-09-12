# 0517 DOCX publication profiling and guard protocol

The unresolved 0500 phase signal is the `p128-k1-owned` batch-versus-scalar
publication interval: `publish_ns` p50 is 15.227% higher (p95 14.791% and
p99 13.888%). The same row's whole-child RSS change is only 1.705%. The
retained greater-than-five-percent RSS flag is `p512-k8-file` (5.770% against
the after scalar control and 7.887% against scalar before). Keep these cases
separate; the 0500 evidence does not establish a K=1 publication RSS
regression.

## Current timing and signals

The source-bound primary case is
`crates/litchi-docx/examples/managed_paragraph_batch_perf.rs`. It accepts:

```
--paragraphs {128|512} --replacements {1|8|32}
--source {owned|file} --mode {repeated|batch}
--warmups 3 --samples 30 --repeats 2
--artifact-dir <fresh-directory> --output <new-file>
```

Fixture construction, the unmanaged and managed preflights, and expected
output generation are before the operation clock. The clock covers
source-backed open, edit staging, commit, `publish_document_commit_to_stream`,
and commit destruction. In this harness `publish_ns` also contains the drop of
the returned published snapshot; output semantic/reopen checks, untouched
media and opaque-member checks, hashing, and CSV writing follow the clock.
Each child therefore emits 66 rows (six warmups and 60 measurements) but
`/usr/bin/time -v` supplies one whole-child maximum RSS value.

The separate `tools/perf-baseline` `docx-managed-edit` selector is useful only
as a diagnostic control. Its `--phase-diagnostics` fields split open, edit,
commit, diagnostics/XML identity, publication, published-snapshot drop, and
commit drop, with publication excluding the snapshot drop. It measures one
paragraph on its 0188 corpus and is not interchangeable with the 0500 batch
fixture. Both sets of phase fields are wall-clock `Instant` intervals, not CPU
time.

The dedicated example records source range calls/bytes, cache counters, sink
write counts, output identity, monotonic `InputBytes`/`Work`, and managed
memory/object release. It has no allocation count, allocator peak, or
publication-local RSS. Its `perf stat` supplement is whole-child only
(`cycles`, `instructions`, `branches`, `branch-misses`, `cache-misses`,
context switches, CPU migrations, and page faults); setup, preflight, all
samples, verification, and report handling are included. These counters must
not be presented as publication CPU attribution. The allocator binary and
phase diagnostics in `tools/perf-baseline` remain full-lifecycle evidence.

## Recommended zero-code CPU profile

Callgrind can provide a bounded publication-function diagnostic without
changing the source or adding phase markers. The harness performs exactly two
preflight publication calls before the measured `run_sample` call. For one
fresh process with `--warmups 0 --samples 1 --repeats 1`, use the frozen
demangled symbols from `nm -C` and run the equivalent of:

```
valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect='<exact publish_document_commit_to_stream symbol>' \
  --zero-before='*managed_paragraph_batch_perf::run_sample' \
  --callgrind-out-file=<profile> <frozen-binary> \
  --paragraphs 128 --replacements 1 --source owned --mode batch \
  --warmups 0 --samples 1 --repeats 1 \
  --artifact-dir <fresh-directory> --output <new-file>
```

Callgrind toggles collection on entry to the selected function and off again
on its exit. Thus each of the two preflight publication calls is a complete
on/off interval; `zero-before=run_sample` clears those counts while collection
is off, and the one measured publication gets one complete on/off interval.
No second toggle on `run_sample` is needed. Keep one measured sample per fresh
process: with multiple samples, a zero-before point on every `run_sample`
entry leaves the final profile dependent on the last sample rather than
providing one unambiguous aggregate. Verify the frozen binary contains the
publication and `run_sample` symbols, the raw post-zero call graph contains
exactly one positive publication call, the fixture/output oracles pass, and
the owned artifact directory is removed. Reject a profile with an absent,
duplicated, ambiguous, or inlined-away selected symbol.

Use `callgrind_annotate --auto=no --threshold=100 --show-percs=no
--tree=both` twice, with `--inclusive=yes` for the selected total and
`--inclusive=no` for the selected self row plus its direct `>` edges. Retain
the raw profile, both annotations, command, binary SHA-256, symbol inventory,
and cleanup receipt. Compare the selected function's inclusive Ir and call count across
matched binaries; total-process Ir after the publication call also includes
the harness's post-publication verification and is not a publication metric.
The selected function ends before the caller drops the returned snapshot, so
its cost cannot equal the native `publish_ns` field, which deliberately
includes that drop. Callgrind Ir is an instrumented instruction-read count,
not hardware cycles or elapsed time. Use two independent fresh profile
processes per arm; do not use this toggle/zero arrangement with three or more
samples in one process without a reset/wrapper strategy.

`analyze_profiles.py` is the reusable raw-proof check. It resolves `fn`/`cfn`
records, selects the positive incoming edge to the publication symbol, lists
the selected function's direct callees, and requires one positive incoming
edge with one call. It also requires the incoming inclusive Ir to equal the
raw `summary` and the selected self cost plus direct-callee costs. For the
existing preflight and 12 profile-r1/profile-r2 lanes, every profile passes
those checks; the preflight selected total is 6,662,867 Ir (self 474, direct
callees 6,662,393). The inclusive and exclusive tree rows agree with the raw
parser for all 13 existing profiles. Re-run the same parser for later profile lanes, for
example:

```
python3 -B docs/performance/results/change-0517/analyze_profiles.py \
  'docs/performance/results/change-0517/profile-r1/*.callgrind' \
  'docs/performance/results/change-0517/profile-r2/*.callgrind' \
  --output docs/performance/results/change-0517/profile-analysis.json
```

The two current profile repeats cover these scoped inclusive-Ir ranges (the
two route values are kept distinct):

| Case | Repeated | Batch |
| --- | ---: | ---: |
| p128 K=1 owned | 6,660,110–6,661,889 | 6,660,412–6,661,707 |
| p512 K=1 owned | 22,472,859–22,473,119 | 22,473,282–22,476,195 |
| p512 K=32 owned | 22,473,401–22,479,305 | 22,484,846–22,485,900 |

These are two one-sample diagnostic repeats on one binary, not a replacement
for the 30-sample native guard.

The completed candidate lanes are matched in
[`profile-comparison.md`](profile-comparison.md) and
[`profile-comparison.json`](profile-comparison.json). All 12 baseline and 12
candidate profiles pass the raw and annotation checks. Candidate publication
inclusive Ir is about 17.9% lower for p128 K=1 and 21.0% lower for p512 K=1
and K=32; the direct topology-owner edge is about 32.3% lower at p128 and
43.7% lower at p512, while inclusive `validate_source_xml` is about 33.3%
lower across the six arms. This is instruction-count attribution for the
duplicate-validation change and remains separate from native timing/RSS
admission.

Profile the owned source first to avoid filesystem noise. The focused arms
should be:

| Profile arm | Purpose |
| --- | --- |
| p128 K=1 repeated and batch | same-binary route control for the small flagged case |
| p512 K=1 repeated and batch | K=1 size control |
| p512 K=32 repeated and batch | large selection/common-publication control |

The 0517 profile lanes use the same current binary for both route choices.
Bind `source`, `mode`, case dimensions, binary hash, and source-manifest hash
in every receipt; keep any future baseline/candidate binary comparison as a
separate same-route comparison.

## Guard and formal matrix

For the 0517 source-bound guard, retain 24 route rows on the same current
binary: 128 and 512 paragraphs, K=1/8/32, owned/file sources, and repeated and
batch modes. Run fresh children with
the existing 3-warmup/30-sample/2-repeat settings, serialized under the
measurement lock and with the same CPU affinity. Run two fresh campaigns in
reverse arm order when host time permits. Preserve every row and every
greater-than-five-percent latency, throughput, or RSS flag; do not replace
the formal matrix with the one-sample profile.

The targeted RSS guard should repeat the K=1 owned/file rows and the retained
`p512-k8-file` row in fresh normal-release children. A one-sample process guard
(`--warmups 0 --samples 1 --repeats 1`) can expose startup/high-water changes,
but remains whole-child RSS. Report median and tail across independent
children if it is used; never call it publication-local peak memory. The
standard 66-row child gives only one maximum RSS observation and should remain
descriptive. A publication-local allocation/RSS claim would require a new
explicit phase allocator/sampler and cannot be inferred from the current
budget gauges or GNU time.

Run a separate `perf stat` capture for the focused and retained guard arms if
the events are available, recording event availability and the exact source
and binary hashes. Keep it whole-child and supplementary. The current
toolchain is the repository-pinned `rust-toolchain.toml` version (historical
0499/0500 evidence used 1.98.1); bind the actual `rustc -Vv`, release flags,
affinity, environment, harness SHA-256, and binary SHA-256 for every campaign
instead of mixing those historical binaries. OLE2/OOXML remain in scope;
ODF and iWork are outside this protocol.
