# 0516 final verifier requirements

This document defines the checks the final 0516 verifier must perform over the
retained OOXML/XLSX evidence. It is a requirements record, not an integrated
verifier and not a performance claim. The existing helpers remain useful
components, but their individual scopes are intentionally narrower than this
check.

The plan fixes the base revision at
`c5abaef129f5ae0000a925b295cf6507b7474e3c`, keeps the `before` control
manifest equal to the 0515 source epoch, and gives priority to OLE2/OOXML.
ODF remains deferred. The final output must retain
`performance_claim: "none"` and `claim_authorized: false`, including when the
candidate is retained. Pilot, profile, and allocator observations are
descriptive until all declared formal gates pass, and an `Ir` profile is
instruction attribution rather than elapsed time.

## Decision contract

The verifier should require one strict `decision.json` object with a schema
version and these semantic fields:

```json
{
  "schema": "litchi-0516-verifier-decision-v1",
  "decision": "reject",
  "final_epoch": "baseline-tests",
  "candidate_epoch": "fifth-unit",
  "baseline_epoch": "before",
  "baseline_test_epoch": "baseline-tests",
  "source": {
    "final_manifest": "baseline-tests/source-manifest.json",
    "candidate_manifest": "fifth-unit/source-manifest.json",
    "baseline_test_manifest": "baseline-tests/source-manifest.json",
    "production_restored": true,
    "baseline_test_fixes_retained": true
  },
  "tests": {
    "required_success_receipts": [],
    "expected_failures": "expected-failures.json",
    "baseline_failures": "baseline-failures.json"
  },
  "builds": {
    "required_receipts": []
  },
  "captures": {
    "required_lanes": [],
    "formal_abba": false
  },
  "performance_claim": "none",
  "claim_authorized": false
}
```

The example is the rejection shape. For `decision: "retain"`,
`final_epoch` must be the final candidate source epoch,
`source.production_restored` must be false, and `captures.formal_abba` must be
true. For `decision: "reject"`, the candidate epoch and all failed evidence
remain authenticated, while `final_epoch` must name the restored source
epoch. If the two independent baseline test repairs are retained, that epoch
is `baseline-tests`; otherwise it is `before`. A decision with an absent or
ambiguous final epoch is incomplete and must fail closed.

The verifier must require the lists in `tests`, `builds`, and `captures` to be
nonempty and to name concrete retained receipt paths. It must reject unknown
keys where the contract is intended to be closed, reject duplicate paths, and
check that every listed path is inside the evidence root. The decision must
state whether after captures and formal ABBA are required; the verifier must
derive the required set from the decision and plan rather than treating an
empty list as “nothing to check.”

## Source epochs and the retained baseline fixes

The source graph currently has a control branch and a candidate branch:

```text
before  -> initial-unit -> second-unit -> third-unit -> fourth-unit
   `-> baseline-tests -> fifth-unit
```

The `previous_epoch` field in `candidate-patch.json` is part of the custody
chain. The existing `custody.py` replays patches but does not check this
relationship; the final verifier must check it for every retained candidate
epoch. Every patch record must bind the plan base revision, patch hash,
manifest hash, changed path set, and candidate file hashes. A replay from
base blobs in a temporary directory must produce exactly the recorded changed
files and hashes, and the temporary directory must disappear afterward.

`baseline-tests` is a separate, test-only repair of the unchanged control.
Its only changed paths must remain:

* `crates/litchi-xlsx/tests/source_backed_cell_values.rs`
* `crates/litchi-xlsx/tests/source_backed_row_visibility.rs`

Its source manifest must differ from `before` only at those two paths. The
fixed suite receipt records 1263 passed, zero failed, zero ignored, and zero
filtered. The final candidate epoch (`fifth-unit` at the time of this
record) must carry those same two fixed file hashes when the fixes are
retained. The verifier must compare the final and candidate manifests against
`baseline-tests`, rather than treating the two test fixes as candidate
production changes or silently running the candidate against the unfixed
control tests.

Before receipts legitimately bind the original `before` manifest, because
they were captured before the independent test repair. Candidate captures and
final test/quality receipts must bind the final candidate manifest (including
the retained test fixes). A rejection may restore production files to the
`before` hashes while retaining the two test fixes; in that case the final
manifest is the `baseline-tests` manifest, the candidate manifest remains an
archived distinct object, and after receipts still bind the candidate epoch
that produced them.

For every source manifest used by a receipt, the verifier must check:

* all paths and lowercase SHA-256 values are valid and the manifest is
  nonempty;
* `before` hashes equal the declared base revision and the 0515 control
  manifest;
* a retained candidate manifest has the expected XLSX-only production change
  set and the expected focused tests;
* the live source equals `final_epoch` when the decision retains or after the
  source is restored; and
* a receipt's `source_manifest_sha256` equals the exact retained manifest
  named or unambiguously selected for that receipt.

Guard and fallback probe manifests are separate from production source
manifests. Their probe sources, canonical allocator observer files, lockfiles,
and executable hashes must be checked independently; a guard receipt must not
be accepted as a production receipt merely because its basename is similar.

## Failed epoch retention

`expected-failures.json` is an allowlist of historical nonzero commands, not a
way to hide missing evidence. For each entry, the verifier must require the
receipt and its log to exist, require the recorded exit code to equal the
declared code, verify `source_unchanged`, source manifest hash, log hash,
command, and any artifact hashes, and preserve the failure reason in the
final inventory. Every retained nonzero receipt must appear in the allowlist;
every allowlisted receipt must be present. A later successful run must not
overwrite or relabel an earlier failed receipt.

The current failed epochs have distinct meanings and must remain distinct:

* `initial-unit`, `second-unit`, `third-unit`, and `fourth-unit` are candidate
  development failures with their source snapshots and patches;
* `fourth-unit/xlsx-features-retry-receipt.json` is still the candidate run
  that exposed the two stale no-op budget tests;
* `before/xlsx-source-backed-values-receipt.json` and
  `before/xlsx-features-known-receipt.json` are unchanged-control failures;
  they are not candidate regressions; and
* `baseline-tests/xlsx-features-receipt.json` is the independent corrected
  control result and must not be substituted for either failed receipt.

`baseline-failures.json` must be checked against both failed receipts and the
fixed suite receipt. The two named control failures must retain the same typed
errors and the corrected suite must be the only evidence used to establish the
repaired test contract. Skipping the tests in a candidate command cannot count
as a successful full suite.

## Required lane set

The plan's row cardinalities and timing settings are part of the final
contract. A report that exists with a missing, duplicate, unknown, or malformed
row is an error. A partial comparison must never be promoted to an admission
result.

| family | expected rows | preflight | pilot | allocator repeats | formal native if retained |
| --- | ---: | ---: | ---: | ---: | ---: |
| main | 3 shapes × 4 cases = 12 | 1 sample / 0 warmups | 20 / 2 | `allocator-r1`, `allocator-r2`, 10 / 1 | `r1`, `r2`, 500 / 5, ABBA |
| guard | 3 shapes × 7 scenarios = 21 | 1 / 0 | 20 / 2 | `allocator-r1`, `allocator-r2`, 10 / 1 | `r1`, `r2`, 100 / 3, ABBA |
| fallback | 3 shapes × 2 scenarios = 6 | 1 / 0 | 20 / 2 | `allocator-r1`, `allocator-r2`, 10 / 1 | `r1`, `r2`, 100 / 3, ABBA |
| profile | dense-wide, one-percent commit-save | 3 samples / 0 warmups | diagnostic | diagnostic | diagnostic only |

For a final **reject** decision after a completed admission attempt, require
both `before` and `after` preflight, pilot, allocator-r1, and allocator-r2
lanes for all three measured families, plus the single diagnostic profile lane.
Formal native ABBA is not required to record a rejection after a failed
pilot/allocator/profile gate, but the output must remain claim-free.

For a final **retain** decision, require that same complete before/after set
plus the formal native lanes in the table. The native receipt start times and
durations must prove the declared serial order
`before/r1, after/r1, after/r2, before/r2` for each formal family. If allocator
data is used in the retention decision, require both stages' r1/r2 allocator
captures with matching rows and the same serial, nonoverlapping operation
rule; do not infer a formal memory comparison from one side.

The before-only profile is valid while the evidence is partial. A final
decision that lists a profile comparison must have both profile artifacts and
matching receipt/annotation hashes. A profile that is present but incomplete
must fail; a missing after profile may only be represented as an explicitly
diagnostic, non-comparative lane.

The verifier must also require the lane order and protocol values recorded by
`plan.json` and `fallback-plan.json`: CPU 2, one worker where applicable,
serial captures, fresh child/process isolation, and no simultaneous local
builds or heavy checks. It must distinguish the normal 12-row main corpus,
the seven-scenario public guard, and the two-scenario x14ac fallback guard.

## Receipt, artifact, and command binding

`audit.py` is a useful receipt inventory, but it currently accepts any
recorded command, skips binary verification when the scratch binary is absent,
and does not enforce lane completeness. The final verifier must add these
checks:

* Every receipt is strict JSON with no duplicate keys or nonfinite numbers;
  `exit_code`, `source_unchanged`, timestamps, positive duration, source
  manifest digest, binary digest where applicable, and artifact map have the
  expected types. Failed receipts use the declared failure contract.
* Every artifact path is a canonical, traversal-free basename in its stage;
  every listed artifact exists, is a regular non-symlink file, is nonempty
  when the producer requires nonempty output, and hashes exactly as recorded.
  The receipt must bind the report, catalog, log, and any raw/profile output;
  the annotation manifest must separately bind inclusive and exclusive text
  to the raw profile.
* Build receipts bind the exact normal or allocator Cargo command, source
  manifest, build log, release binary name, and binary hash. Guard/fallback
  build receipts additionally bind the independent probe manifest and offline
  lockfile. A post-cleanup verifier may authenticate a binary by its retained
  hash without requiring the scratch copy, but it must not silently skip the
  hash check while scratch still exists.
* Native capture commands bind `/usr/bin/time -v`, `taskset -c 2`, the correct
  stage binary and hash, exact case and shape lists, samples/warmups, report
  and catalog basenames, and the lane-specific output paths. The allocator
  command must select the allocator binary and must never supply its elapsed
  or RSS as normal latency evidence.
* Profile commands bind Valgrind Callgrind, `--collect-atstart=no`, the exact
  `xlsx_commit_save_operation` toggle, dense-wide one-percent commit-save,
  three samples and zero warmups, and the raw output path. The profile log's
  `brk segment overflow` status is retained as a caveat only.
* Guard/fallback capture commands bind the right probe, CPU, shapes, samples,
  warmups, report path, probe manifest, and allocator feature. A fallback
  report must not be accepted as a plain guard report.
* Quality/test receipts bind the exact command from `check.py`, including
  package, feature, test target, skips, and warning policy. A skipped test
  receipt cannot satisfy a required full-suite receipt.

At pre-cleanup time, available binaries and source files must hash to their
receipts. After cleanup, the final verifier must use retained hashes and
manifest/patch replay instead of silently accepting missing paths. Relevant
environment fields (`CARGO_TARGET_DIR`, jobs, incremental/debug settings,
`TMPDIR`, `RUSTDOCFLAGS`, and inherited Rust/loader variables) must be
consistent with the recorded command and host record.

## Correctness and report validation

Compose `correctness.py` with the standalone 0514 verifier helpers instead of
duplicating their row parser. `correctness.py` deliberately walks only reports
that exist, so the final verifier must first enforce the required lane set and
then use its checks. For each report, require the expected schema/tool/binary
identity, environment and configuration, catalog binding, exact corpus and
output identity, sink counters, operation alignment, and every raw elapsed and
allocation vector. For guard and fallback rows also require the public oracle,
source-marker, x14ac/default-descent, and untouched/changed-cell checks.

Matched before/after rows must use the same shape/case or shape/scenario key,
corpus identity, sink identity, output hash, and binary/source role declared by
the protocol. Do not compare only rows that happen to overlap. Require the
same-build before repeat identities and retain their drift flags separately
from candidate deltas.

Use `metrics.py` to expose every individual row and recompute p50, mean, p95,
and p99 from raw vectors. Its RSS reader is correctly limited to one GNU-time
line from normal logs and excludes allocator lanes; a final verifier must make
missing required logs/RSS an error rather than allowing `_rss` to skip them.
Its descriptive comparisons may include an incomplete capture, so they cannot
by themselves establish completeness or admission. Allocation counters and
peak fields must stay in allocator scope, with incremental region peak kept
separate from absolute peak and no allocator elapsed/RSS comparison.

Use `profiles.py` for the diagnostic profile. It dynamically reuses the 0515
raw parser and validates the current no-percent annotation format, the raw to
inclusive selected-edge `Ir` equality, exactly three calls for runner → helper,
helper → `Edit::commit`, and helper → `PackageWriter::write_to_stream`, plus
exclusive attribution and the warning status. The final verifier must add the
profile receipt command, source/binary identity, and stage binding that this
focused extractor intentionally leaves outside its scope.

## Negative vectors and fail-closed behavior

The final verification run must exercise rejection paths in memory and retain
their boolean results. At minimum:

* remove one elapsed sample and alter one elapsed sample while preserving the
  old statistics; `_validate_elapsed` must reject both;
* remove one operation sample index or reorder it; operation alignment must be
  rejected;
* remove a main, guard, or fallback row; duplicate a row key; add an unknown
  row; and corrupt a row's shape/scenario; completeness must be rejected;
* shorten or reorder allocator vectors, publish allocator values on a normal
  row, make region peak precede `live_bytes_before`/`live_bytes_after`, or make
  it exceed `peak_live_bytes_after`; all must be rejected;
* alter a receipt's source, binary, command, stage, artifact, or annotation
  digest; alter an annotation's raw digest; or remove a required log/RSS line;
  custody must be rejected;
* alter a candidate patch path, base revision, previous epoch, changed-file
  inventory, or candidate hash; isolated patch replay must be rejected; and
* present only a partial after lane or a profile report without its receipt;
  the final decision must be incomplete rather than silently downgraded to a
  descriptive pass.

These mutations should target deep copies or temporary evidence copies and
must never modify retained evidence. A negative vector is useful only when the
same parser/helper that validates the real report rejects it; checking a
hard-coded `True` in a receipt is insufficient.

## Cleanup and post-cleanup replay

Before cleanup, record a complete verifier result and an inventory/hash of all
retained evidence. Cleanup may remove only the owned scratch worktree,
executables, target caches, and generated task caches named by the cleanup
record. It must retain source manifests, candidate patches and metadata,
receipts, logs, reports, catalogs, raw profiles, annotations, review records,
negative-vector results, and the final decision.

After cleanup, run the verifier from a copied or relocated evidence bundle
with the owned scratch paths absent. The replay must:

1. authenticate every retained receipt and manifest without relying on an
   absolute `/tmp` path or a live candidate binary;
2. replay each candidate patch from base Git blobs in an isolated temporary
   directory and verify its final file hashes and removal;
3. validate the final live source against `final_epoch` (or the restored
   baseline/test-fix manifest for a rejection);
4. produce the same canonical decision, lane-completeness, custody, and
   descriptive-metric result as the pre-cleanup run; and
5. exit zero with no stderr, with the cleanup record and retained inventory
   hashes unchanged.

`custody.py` already supplies the useful temporary patch replay primitive, but
it does not check final-source selection, epoch chaining, or copied-bundle
verifier replay. Those checks belong in the final verifier. A missing scratch
binary after cleanup is expected only when its pre-cleanup receipt and hash
were authenticated and the cleanup record accounts for its removal.

The final result must expose any failed epochs, drift/adverse flags, profile
warnings, missing or excluded measurements, and whether formal ABBA was
completed. It must never turn a lower descriptive latency or `Ir` value into a
speedup claim, and it must not advance the ODF lane before the OLE2/OOXML goal
is complete.
