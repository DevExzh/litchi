# Change 0168: XLS numeric validation fusion

Date: 2026-08-17

Status: production work elimination retained; release latency claim withheld.
The clean release comparison found lower complete-workflow and semantic-commit
p50, mean, p95, and p99 values in both paired directions for Number and
RK/MulRK. Same-implementation temporal drift nevertheless exceeded the
predeclared 5% gate, so those distributions remain descriptive rather than an
acceptance-grade speedup or tail-latency result.

## Mechanism and validation boundary

The native XLS plan-only fixed-width numeric path already asked the common CFB
splice planner to fingerprint the source, construct a positional composed
target, reopen that target as CFB, verify every selected range, and fingerprint
the source and composed target again. XLS then called
`ValidatedOverlayPlan::composed_source()` twice around its independent BIFF
semantic validation. Each call repeated a complete source scan while hashing
both source and composed-target bytes.

CFB now exposes an additive low-level owner-validation seam. It passes the
exact `ComposedOverlaySource` used by the CFB reopen to a synchronous format-
owner callback after structural/range validation and before the final complete
fingerprint fence. The callback cannot access physical spans or publish bytes.
Its generic error boundary preserves native XLS semantic errors while CFB
errors retain their existing typed conversion. Exact byte no-ops skip the
callback and return no owner result.

The XLS owner validates that same view through `Workbook::new`, retaining
worksheet coverage, workbook and worksheet protection, macro refusal, and
independent numeric public readback. It returns the composed view's target
version from inside the fence. The change therefore removes only the two
post-plan `composed_source()` calls. Initial/final source and target
fingerprints, CFB reopen, selected-range comparison, source preconditions,
signed/encrypted/DRM/macro/protection policy, immutable plan, partial-sink
typing, atomic save, and publication preflight remain.

The removed work is deterministic from the implementation and
`FINGERPRINT_CHUNK_BYTES = 1 MiB`:

| Family | CFB bytes | complete scans removed per sample | source bytes no longer scanned | `ReadAt` calls no longer issued |
|---|---:|---:|---:|---:|
| Number | 16,995,840 | 2 | 33,991,680 | 34 |
| RK/MulRK | 202,752 | 2 | 405,504 | 2 |

These are logical in-memory source reads derived from the code path, not
measured physical I/O. The harness's `source.read_*` vectors describe owned-
source ingress and intentionally do not count the internal fingerprint reads.

## Verification

The frozen production revision passed:

- 225 `litchi-cfb` unit tests and all CFB integration/example targets;
- 127 `litchi-ole-common` unit tests and all OLE-common integration/example
  targets;
- all 1,015 `litchi-xls` library tests;
- every other XLS all-target test after skipping
  `all_profiles_round_trip_and_emit_exact_filepass_families`;
- strict production Clippy for CFB, OLE-common, and XLS with
  `-D warnings -D deprecated`;
- formatting and diff checks;
- two independent read-only reviews of the API boundary and adversarial
  fingerprint/security behavior.

The skipped XLS writer-encryption test also fails at the clean control revision
with the same expected-record mismatch. It is therefore a pre-existing,
unrelated all-target failure rather than a regression in the numeric
source-backed path.

New CFB regressions prove that the owner callback observes replacement and
untouched bytes through the exact composed view, byte no-ops omit the callback,
native owner errors return no plan, and a stable-version source mutation during
owner validation is rejected by the final complete fingerprint fence.

## Clean release protocol

Control `66ad0e76c4b2b587c890999fd1b0b73daafd42e7` and candidate
`3b4666ffe3c141d73e4f2bd53a7cd24eea6f3e7a` were built from clean detached
worktrees. Their binary SHA-256 values are respectively
`e4fcf89c82d800e93ce55b775e52c98781e2d3f7047a33e15b51b2d4caefb6e2`
and
`43d659db23738e07a9d6561be94d1e966a7590b5811f85bd56487350e56098bb`.
Every raw record reports `git_worktree_dirty: false`, CPU affinity `2`, one
available logical CPU, Rust 1.95.0, the Rust system allocator, and the AMD EPYC
9575F host.

Fresh processes ran strictly `A1 control, B1 candidate, B2 candidate, A2
control`. Each family used 20 warmups and 500 retained samples of the same
plan-only selector on both revisions. Number changes one eight-byte value in a
16,995,840-byte opaque-heavy CFB. RK/MulRK changes three four-byte values in a
202,752-byte CFB. p50 uses the integer midpoint of the two central sorted
values; p95 and p99 use nearest rank.

All 4,000 retained samples preserve the exact source/output hashes, CFB and
Workbook lengths, sink topology, family-specific splice counts and bytes,
source/target fingerprints, semantic reopen and readback, opaque streams, and
the runner's no-op, stale/foreign, partial-sink, signed, encrypted, macro,
protection, and real-producer gates. Complete target materialization at commit
remains zero for every sample.

## Descriptive result and rejection boundary

Positive values mean the candidate was lower than the control. `total_ns` is
the complete edit, semantic commit, and publication interval retained by the
XLS numeric runner. `commit_ns` is the plan construction and semantic target-
validation interval where the two scans were removed.

| Family | Phase | Pair | p50 | mean | p95 | p99 |
|---|---|---|---:|---:|---:|---:|
| Number | complete workflow | A1 -> B1 | 27.15% | 26.51% | 21.99% | 19.72% |
| Number | complete workflow | B2 -> A2 | 27.32% | 26.92% | 26.26% | 26.30% |
| Number | semantic commit | A1 -> B1 | 47.76% | 47.36% | 43.98% | 42.28% |
| Number | semantic commit | B2 -> A2 | 48.04% | 47.81% | 47.39% | 47.62% |
| RK/MulRK | complete workflow | A1 -> B1 | 19.22% | 20.21% | 21.44% | 19.87% |
| RK/MulRK | complete workflow | B2 -> A2 | 27.24% | 26.13% | 27.31% | 28.16% |
| RK/MulRK | semantic commit | A1 -> B1 | 40.58% | 41.12% | 39.25% | 37.58% |
| RK/MulRK | semantic commit | B2 -> A2 | 46.03% | 45.22% | 44.74% | 43.80% |

Direction agrees for every listed family, phase, pair, and statistic. The 5%
same-implementation gate fails, however: maximum absolute control drift is
10.56% for RK/MulRK complete-workflow p50, and maximum absolute candidate drift
is 9.81% for RK/MulRK semantic-commit p99. Number control drift also reaches
8.30% at semantic-commit p99. The production work elimination is retained, but
no acceptance-grade latency, tail-latency, allocation, RSS, peak-memory,
physical-I/O, cold-cache, decompression, recompression, or real-producer
improvement is claimed.

Publication code is unchanged and its descriptive directions do not
consistently agree. `/usr/bin/time -v` maximum RSS in A1/B1/B2/A2 order is
194,456/194,200/194,276/194,292 KiB for Number and
159,768/143,408/143,408/143,284 KiB for RK/MulRK. These whole-process sidecars
include corpus construction, warmups, untimed correctness gates, and all
samples; they are retained for reproducibility only.

## Artifacts and reproduction

The [summary](../results/xls-numeric-validation-fusion-0168-summary.json),
[primary statistics](../results/xls-numeric-validation-fusion-0168-primary-stats.tsv),
[comparisons](../results/xls-numeric-validation-fusion-0168-comparisons.tsv),
and [manifest](../results/xls-numeric-validation-fusion-0168-manifest.json)
bind eight compressed raw schema-1 JSON vectors and eight
`/usr/bin/time -v` sidecars.

```sh
taskset -c 2 litchi-perf-baseline \
  --case xls_numeric_plan_only_number_edit_save \
  --warmup 20 --samples 500 --workers 1 --json RESULT.json

taskset -c 2 litchi-perf-baseline \
  --case xls_numeric_plan_only_rk_mulrk_edit_save \
  --warmup 20 --samples 500 --workers 1 --json RESULT.json
```

This change adds no selector and leaves the selectable matrix at 315 names and
the historical default tranche at 36 cases / 198 records. Full BIFF semantic
budget/cancellation accounting, cold filesystem/device evidence, allocator and
peak-memory attribution, broad producer coverage, structural numeric edits,
and the unrelated writer-encryption regression remain open.
