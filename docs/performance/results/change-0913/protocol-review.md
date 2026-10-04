# 0913 matched binary measurement protocol review

Status: **conditional pass pending capture and replay**. This is a read-only review of the
0913 custody and measurement scripts; no build, workload, or profiler capture
was run by this reviewer.

The intended comparison is sound: retain the unchanged 0910 native executable
as the baseline, reconstruct the frozen 0910 source, apply only the bounded
`office.rs`/`payload.rs`/test overlays to an isolated candidate source, build
the candidate with the 0910 release settings, and run both binaries in a
route-specific AB/BA schedule. Each process is pinned to CPU 12, performs 30
warmups and 20 measured samples, and records whole-process `/usr/bin/time`
metrics. Twelve paired processes per route provide 12 baseline/candidate
pairs for both `materialized` and `owned`; ratios must be baseline over
candidate within one route. This design does not authorize a
materialized-versus-owned claim or a claim about allocation, native cycles, or
all Office workloads.

The executed source freeze is appropriately bounded: `prepare.py`
reconstructs the 9355-file 0910 source, and the coordinator invoked `freeze`
with all four explicit overlay paths (`office.rs`, `payload.rs`, the ZIP
differential test, and the OPC concurrency test). The resulting
`source-v1.json` records those four paths, an isolated-index binary patch, and
the full candidate inventory; the plan was written before build or capture.
The benchmark binary and all semantic, source, output, sink, and untouched
member guards remain the admission authority. The quality script also keeps
the known baseline rustfmt debt scoped rather than reformatting unrelated
files.

The quality retry custody is coherent: the initial stale-lock and interrupted
attempts, followed by the offline cache miss in v3, remain as failed receipts;
v4 uses the fixed v3 workspace lock with `--locked`. The v3 seed is
byte-identical to the benchmark lock and records no dropped seed registry
tuples. Fixture removal verifies that the updated quality lock was active
during quality, then restores the frozen source lock and compares the complete
source inventory to `source-v1.json` before capture. Replay checks these prior
attempts and the restoration receipt rather than treating the v4 run in
isolation.

Quality-v4 reached the test compile but failed because the frozen source did
not contain the two archived ZIP corpus members required by
`directory_prefill.rs`: `zip64-central-local-signed.zip` and
`many-small-central-local-signed.zip`. The new `extra-fixtures.py` admission
helper identifies all 13 archived corpus files and records their hashes. The
retry now removes those rows through `fixtures.py remove` and rechecks the
complete frozen inventory alongside the Cargo.lock restoration; replay must
also validate the extra-fixture receipt and removal count before capture is
allowed.

Quality-v5 then failed during the same test command because the reconstructed
source omitted
`docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json`,
which is included by the existing
`crates/litchi-xlsx/tests/form_control_scalar_lifecycle.rs` test. This is a
0910 source-base omission: the archive commit `468fe293…` contains the
11,467-byte corpus, but the 0910 `source-base.json` archive paths exclude
`docs/report`. It is therefore not a 0913 production or test overlay. A
retry must admit that single file with its source-commit/archive digest,
record it in the fixture receipt, remove it before the exact `source-v1`
inventory check, and retain the failed v5 receipt. Until that custody step and
the full quality sequence pass, no quality or capture gate is satisfied.

The v6 retry admits that corpus and the second missing XLSX input used by
`drawing_svg_read.rs`, `tdf169496_hidden_graphic.xlsx`. The latter is an
ignored local fixture, so `more-fixtures.json` retains a copied quality-input
artifact and its digest for provenance after the temporary source copy is
removed. The final replay must bind every v6 quality receipt to
`more-fixtures.json` as well as the earlier fixture manifests; checking the
fixture rows only during cleanup is insufficient custody for a quality run.

V6's test run then reached the DOCX drawing suite and failed because the
reconstructed source also lacks `svg.docx`. That suite uses the paired
`tdf164835_nonDummyLineHeight.docx` input, and the PPTX SVG suite uses
`tdf163852.pptx`, `tdf164622.pptx`, and
`tdf169496_hidden_graphic.pptx`; all are ignored local `3rdparty` inputs
outside the frozen archive. The next fixture admission must record all five
with byte hashes and retained provenance, then remove all temporary copies
before the source-v1 inventory check. A v6 test failure remains retained even
if the subsequent admission succeeds.

The v7 retry admits those five runtime SVG inputs in `runtime-fixtures.json`
and binds the manifest from the quality helper. Replay must validate that
manifest and its retained quality-input hashes, and must check the
`runtime_removed` count in `fixtures-removed.json`; the fixture remover alone
does not make the final packet auditable.

V7 reached the supplemental ZIP adversarial test and failed three of its four
cases at the test's blanket `materialized.is_content_fault()` assertion.
Those malformed and size cases compare equal between the generic and direct
routes, but the generic route exposes the content error through an outer
`ErrorKind::IO` wrapper, which the public predicate intentionally excludes.
The current test revision removes that blanket assertion and retains the
family/diagnostic equality, but it still needs structural checks of the
nested typed fields for checksum and size refusals before this gate is closed.
Changing the production predicate would weaken its transport versus-content
contract. The failed v7 receipt and its v1 test source remain part of the
custody chain.

## Custody notes and required corrections

The retained-binary receipt defect in the first helper revision is corrected:
`build()` now writes `build-<attempt>/binary.json` with the original and
retained descriptors, and `capture()` checks the retained digest. That receipt
must remain create-only and be included in the final custody chain.

The actual capture entry point is `capture.py`, which does not pre-create the
per-leg directory; `run.execute()` creates each directory exactly once. Its
child receipts bind `capture.py` as the capture helper and `run.py` as the
build helper, and the current reducer checks those bindings. The unused
`run.py capture` branch still pre-creates the directory before calling
`execute()` and would fail if invoked directly; it is outside the executed
packet and must not be used for replay.

The executed freeze used four explicit overlays even though the helper's
default tuple still omits the ZIP differential test. That default is a replay
hazard and must be documented in any rerun command; it does not change the
already bound four-overlay manifest. The freeze path also overwrites an
existing overlay in the reconstructed source before writing its exclusive
manifest/patch outputs. A second freeze is therefore forbidden: the source
root and packet must be treated as frozen custody, and any replay must start
from newly prepared marked roots with the same four explicit overlays.

`analyze.py` now supplies the required fail-closed reducer, and its
cross-child guard compares source/output byte counts and SHA-256 while
retaining each leg's path in custody. The earlier path-inclusive comparison
was corrected before capture.

The reducer now validates both sides of the retained-binary handoff. It checks
the build command's original and retained descriptors, verifies the retained
file, requires `build-<attempt>/binary.json` to equal those exact descriptors,
and records that sidecar in the final analysis custody. `capture.py` performs
the corresponding retained-descriptor check before launching children. The
sidecar must remain create-only and retained through replay; a missing or
altered handoff rejects the packet.

The capture helper now satisfies the reducer's environment gate by running
children with the explicit `PATH=/usr/bin:/bin`, locale, and timezone values
and recording `environment_is_complete: true`. The final packet must retain
those receipt fields and verify that no ambient build, profiler, allocator,
debugger, or runtime setting differs between variants.

These are custody requirements, not timing observations. The corrected receipt
chain must bind the baseline and candidate executable path,
byte length, SHA-256, command, source manifest, plan, exit status,
stdout/stderr, process metrics, and every artifact.

## Required capture and reduction gates

After the script correction, the packet should prove all of the following:

* The baseline executable is the retained 0910 binary with its recorded
  digest; the candidate executable is copied once after the candidate build
  and is byte-bound before capture.
* The reconstructed source inventory and candidate overlay manifest remain
  byte-stable through build and capture. Capture should recheck the manifest
  before starting the first child and retain the exact environment used by the
  child processes.
* The frozen plan contains 12 pairs per route, six baseline-first and six
  candidate-first, with every child serial and globally ordered. Each child
  has exactly the declared route, warmup count, sample count, CPU, artifact
  directory, and process-metric output.
* Every child exits zero and its benchmark report proves the same corpus,
  source identity/immutability, semantic output, sink bounds, and preserved
  raw ZIP member census as the 0910 control. A failed or incomplete leg is
  retained and rejects the comparison; it must not be silently dropped.
* The reducer checks all 48 child receipts, cardinalities, ordered sample
  indices, finite phase values, process metrics, and source/output/sink
  digests before calculating per-route distributions. Pair medians are the
  unit for paired ratios and bootstrap resampling; pooled p50/p95/p99 are
  descriptive only. The seed, draw count, schedule, and confidence interval
  method must be retained.
* The timer boundary is unchanged from 0910: source clone, ingress, snapshot,
  edits/commit, apply/serialization, preallocated sink, and route-local
  destruction remain inside the timed body; semantic guards, digest checks,
  artifact writes, and replay remain outside it. Phase clocks are timing
  guards and do not prove a native cost share.

The capture environment is now explicit and minimal. The final packet still
needs to retain the complete child environment receipt and demonstrate that
no build, profiler, allocator, debugger, or ambient runtime setting changes
between baseline and candidate; this is the retained 0910 discipline.

The packet must retain a replayable analysis script and a final review binding
each report, receipt, executable, source manifest, and raw artifact by bytes
and SHA-256. Temporary source/build/capture roots may be removed only after a
before/after replay and an active-cwd check. Cleanup must leave the retained
binary, generated PPTX evidence, and tracked decoded receipts available, and
must not remove unrelated worktrees or the shared iWork conflict state.

No 0913 timing result or copy-removal speedup claim is admissible until the
frozen packet is replayed fail-closed and the complete matched capture passes
these gates.
