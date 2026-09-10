# 0497 measurement method review

This is an independent read-only review of the current 0497 measurement
driver, its focused tests, and the production publication schema. I ran no
Cargo command, build, capture, or analysis. The protected
`/home/zhuhe/code/litchi-spec-gaps` worktree was not accessed.

The reviewed source inputs are bound as follows:

- `measure.py`: `6c7d29f5c2e7fcaf7ea7478c0e54326a3e0f7307e07b77b726f7d683273c930c`
- `test_measure.py`: `746eecb04d1eed8c0b3e19565e9f1624bf0023dfafe64165c19cdfedba5e3873`
- `profile.py`: `b433b8831bfaa3470bd3e67e420bf780a8b8e490152c75366cb89d8fadc5509a`
- `test_profile.py`: `4961dd4516cfc3c142acffa408246b098dc42c8b8e3111d02212b2008c557bee`
- `build.py`: `f2b14181973fb84c56af1e60345b71c7c91ab49f5d4cdc2c1fbb3e5f2db4f529`
- `resume.py`: `984b25a0e838d3441c5b772ff525d13c137d2e65f71c1a68024e3dbe7d69bd01`
- `test_resume.py`: `907e521fbd46bf9be7091ef5e8c3835e0e5fcfcc9d34d09c41317ab52f795785`
- `tools/perf-baseline/src/docx_replayable_tail_append.rs`:
  `b6303b7b353415a4bad997951f98b2dc1c655f795233e10c908f17c684f9b5be`
- `crates/litchi-docx/src/source_backed/tail_append_stream.rs`:
  `963373a9d3b826b5f6651067f6f1e12cc4ad00371b29a58ce35d06d6835c8bf3`

`python3 -B docs/performance/results/change-0497/test_measure.py` passes all
15 focused tests. Independent import-time checks also pass: the formal
inventory contains 288 children and 8,640 measured samples, and the pilot
inventory contains 72 normal-role children and 216 pilot samples. The pilot
is explicitly separate from formal analysis.

The optional syscall profile helper is also internally consistent: `python3
-B docs/performance/results/change-0497/test_profile.py` passes all 8 focused
tests. The warmup-path fix is covered by those tests: atomic traces account
for warmup destinations before the measured destination sequence, preserve
the measured order, and reject extra, mismatched, or incorrectly ordered
write/sync events. Counting traces reject atomic destination paths. The
profile parser retains the complete selected trace and checks sibling write,
data-sync, rename, and parent-sync ordering, so the profile can be used as a
whole-child syscall receipt. Its declared scope includes setup, corpus and
preflight work, warmups, measured children, post-timer oracles, cleanup, and
report serialization; it makes no operation-attribution, latency,
throughput, or optimization claim.

The profile remains a diagnostic receipt rather than a formal matrix result.
Its verifier rehashes the helper, executable, retained build receipt, strace,
report, raw trace, terminal receipt, and cleanup receipt, and checks terminal
chronology and private-directory removal. The current profile binding does
not independently validate the source manifest and gate fields inside the
retained build receipt, and report validation does not reconcile every CLI
run field (source/authored counts, chunk/text settings, replay ceiling,
input mode, and related fields) with the report case. The final custody
bundle must therefore bind a profile to the already validated formal build
and exact command metadata, or retain it explicitly as diagnostic-only. If
the optional profiler cannot be located, that is tool-unavailable/incomplete
evidence; it must not be recorded as a failed benchmark child. Multiple
profile cases likewise require explicit binding before they can support a
single-run claim.

The formal order is now the requested balanced ABBA schedule. The sealed
0489 arm order supplies 18 arms and the two roles supply 36 arm/role pairs:

- ordinals 0--71: repeat 1, each pair `before/hash` followed by
  `after/hash`;
- ordinals 72--143: repeat 2, reversed pair order, each pair
  `after/hash` followed by `before/hash`;
- ordinals 144--215: repeat 1 after-only capability block, each pair
  `counting` followed by `atomic`;
- ordinals 216--287: repeat 2 after-only capability block, reversed pair
  order, each pair `atomic` followed by `counting`.

The driver asserts these positions, route counts, and unique labels while
constructing the inventory. `_collect_lane` requires exactly the frozen
label set, validates each started and terminal run against its protocol
specification, sorts by actual start time, and rejects any start before the
previous terminal time. `_paired_comparisons` forms before/after pairs only
for the two hashing repeats; `_repeat_visibility` retains repeat variance
for every phase and publication route. Counting and atomic routes therefore
remain after-only capabilities and cannot become an invented before/after
comparison.

The timed boundary agrees with the Rust harness and its emitted scope
strings. Fixture construction, store preflight, input preparation, and
atomic private-parent preparation are outside the clock. The clock starts
before source opening/package admission and includes preparation, the
selected production publication, and publication drop. The atomic interval
includes the sibling write, data sync, replacement, and parent-directory
sync. Readback, semantic/raw-member checks, inverse fixture checks, replay
cleanup, report serialization, and process/allocation endpoint snapshots
are outside the elapsed timestamp. The command receipt is built from the
sealed `/usr/bin/time` and CPU-2 `taskset` command, and the shared lock
serializes the complete lane.

The default hashing report continues through the sealed 0489/0484 validator.
Counting reports retain the raw production accepted-byte, write-call,
largest-write, and histogram observations; their artifact digest comes from
the production `ParagraphStreamPublication` proof rather than a synthetic
sink digest. The temporary legacy projection removes only the publication
extension and supplies that proof to the sealed validator. Atomic reports
require the exact publication proof, cleaned private destination, output
byte/hash identity, and complete post-timer oracle. Atomic sink option fields
are accepted only as an empty legacy shape or all-null serialized options;
fabricated write counts or digests are rejected.

The previous atomic configuration metadata gap is resolved in the current
source. `_validate_new_report_config` builds and compares the complete
selected-arm contract, including provider and authored provider, counts and
modes, lifecycle, replay opens and ceiling, replay directory and sync,
compression, input mode/backing/identity, range/delay/overhead/rate fields,
sink contract, fixture path, and publication. The case identity checks bind
the corresponding provider, corpus, replay ceiling, compression, input
identity, and sink fields. Terminal argv and private input/replay paths are
also checked against the same frozen arm.

Source read counters retain request and return totals plus histograms for
every sample; the sealed checks require positive totals, returned bytes no
greater than requested bytes, and histogram conservation. Authored and
replay counters are checked against the sealed event and provider proofs.
Allocator reports preserve every raw counter, enforce zero-net live-byte and
allocation-byte conservation, and bound both process and region peaks. The
derived operation increment is explicitly `region_peak_live_bytes -
live_bytes_before`; normal-role allocator vectors are unavailable rather
than fabricated. GNU-time RSS is represented as one whole-child observation
(`n: 1`) and is not treated as 30 independent operation samples.

Analysis uses only default hashing for before/after comparisons. It keeps
both repeat-block deltas and adverse individual repeats, uses deterministic
within-child resampling for sample vectors, and uses the paired bootstrap
only over the two repeat-block deltas. It does not pair independent child
sample vectors. Counting and atomic observations remain capability records;
the claims explicitly disable an atomic speedup assertion. Raw publication,
allocator, RSS, process, and source/replay evidence remain available for
later analysis.

The build and provenance receipts are consolidated under these exact schemas.
`builds.json` has schema
`docx-replayable-tail-publication-builds-v1` and exactly the four keys
`before/normal`, `before/allocator`, `after/normal`, and `after/allocator`.
Each record must contain only `binary`, `git_revision`, `source_manifest`,
and `gate`; each metadata object must retain the exact `path`, `bytes`, and
`sha256` fields. Each gate must contain exactly `argv`, `cwd`, `environment`,
`source_manifest`, `started_ns`, `driver`, `pid`, `exit_code`, `finished_ns`,
`source_unchanged`, `stdout`, and `stderr`, and must be a successful
locked-offline release build with the expected build environment, matching
source manifest, unchanged source, and retained stdout/stderr artifacts. CPU
2 is enforced by the capture command; it is not a compilation-environment
requirement. The shared `HARNESS_CUSTODY_FILES` substrate hashes must match
between before and after; the full manifests may differ in the six-file
reviewed candidate allowlist, and no other difference is admissible. The
formal loader currently checks only the basename of the gate driver
(`build.py`), so the final provenance must also bind the exact current
`build.py` path and hash rather than relying on that basename check.

`provenance.json` should use the established `docx-phase-provenance-v1`
schema with exactly `schema`, `inputs`, `production_change`, and
`historical_evidence` at top level. `inputs` must contain nonempty exact
path/size/hash metadata for every source, helper, build, protocol, fixture,
and retained evidence input that the final bundle relies on. The frozen
protocol must bind the complete provenance file, all four build records, their
source manifests, and the current hashes above; any later source, helper,
binary, gate, report, or terminal change must fail the revalidation. The
validated top-level `builds.json`, `provenance.json`, and `protocol.json`
receipts now exist. The pilot has verified terminal custody; the formal
capture stopped at the recorded ENOSPC interruption and remains the
outstanding evidence needed for formal result acceptance.

On the current read-only `formal1` snapshot, all 144 legacy hashing children
have passing terminal receipts. Re-running the frozen report and terminal
validators through a unique module import passes every completed legacy child;
their before/after source, authored, sink-counter, output-hash, oracle, and
allocator-conservation fields agree pairwise. This is a correctness and
custody check, not a completed performance analysis. The provisional timing
audit has three individual median deltas above the 5% adverse threshold:
repeat 1 normal `file_store-owned-s64-a64` (+5.22%), repeat 1 allocator the
same arm (+7.63%), and repeat 2 normal `file_store-owned-s64-a16384`
(+6.30%). Across the 18 legacy arms, normal-role medians are generally about
1% higher after while allocator medians are centered near zero; the final
analysis must retain these flags and repeat-block variance rather than
describing the legacy route as uniformly unchanged.

The retained `capture-cleanup-overlap.json` records emergency shared-disk
cleanup overlapping both `r2-before` and `r2-after` allocator latency
`s64-a16384` hashing children for approximately 1.20 seconds and 1.04
seconds, respectively. The receipt says the interval covers whole-child
execution and cannot assign writeback interference to individual timed
samples. That pair must therefore be marked shared-host-confounded or
excluded from any causal before/after interpretation; it does not invalidate
the report or oracle receipts themselves.

The formal coordinator then stopped at ordinal 191,
`r1-after-atomic_path-allocator-deterministic-latency-s131072-a64-short-c64`,
while writing `replay-cleanup.json`. The raw directory contains only the
running `started.json` and zero-byte stdout, stderr, resource, and cleanup
files; it has no report or terminal receipt. Its private scratch contains
only the empty `tmp` directory. `resume.py --check` independently validates
the passing ordinals 0--190, their frozen chronology, and the exact 97-run
suffix; `test_resume.py` passes 11 focused safety tests. The continuation
driver archives these raw files as `.raw` names under
`interrupted/formal1-enospc`, records the original paths and hashes plus the
unknown child exit scope, and never fabricates a terminal receipt. It then
launches only the exact frozen suffix under the shared CPU lock, with
immutable resume start/terminal receipts carrying timestamps, coordinator
exit status, and input hashes. The final collector still requires the full
288-label formal namespace and exact chronology, so the archived interruption
cannot silently become a completed child or support a result claim.

At this review point the frozen protocol and build/source bindings validate,
but formal capture and analysis are not complete, so this method review makes
no performance or route-result claim. Formal acceptance still needs
successful formal terminal custody and chronology checks, retained raw
reports, conservation and oracle validation, cleanup receipts, and analysis
regenerated from those raw receipts.
