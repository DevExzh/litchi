# 0894 selected metadata paired qualification

This packet defines the bounded paired experiment for the selected worksheet
metadata route. It does not reuse the 0891 executable: both baseline and
candidate must be fresh release builds from the 506aa2eda source family, with
the same post-endpoint-per-worksheet-tolerance-v1 harness guard. The retained
0891 and 0892 binaries remain historical evidence only.

The packet owns the capture protocol and semantic/custody helpers. The source
and build helper owns the baseline archive, candidate overlay, disposable
source/build roots, and four retained executable descriptors. No helper in
this packet creates a worktree or changes production source.

## Current result and recovery status

The completed [quality-v9 receipt](attempts/quality-v9/quality-receipt.json),
[corpus comparison](corpus-v1/comparison-v2.json),
[reduction](reduction.json), and [independent verification](verification.json)
record a semantic match and a measured rejection of the candidate's frozen
primary gate: the producer XLSX full-cell p50 candidate-over-baseline ratio is
1.8972120877 (95% interval 1.8743111450–1.9204038076). The
[restoration receipt](restoration.json) records restoration of three tracked
candidate files (one implementation and two test files) and removal of the two
candidate additions.

[Cleanup-v2](cleanup.json) passes after accepting the previously
hash-verified early quality-root removal; the original cleanup failure remains
retained in [cleanup-failure.json](cleanup-failure.json). The
[verification-after-cleanup receipt](verification-after-cleanup.json) passes.
[Capture review](capture-review.json), [corpus replay after cleanup](corpus-v1/verification-after-cleanup.json),
and [final review](final-review.json) bind the completed evidence. The
[decision](decision.json) rejects the candidate; production is restored and
all owned temporary source/build roots are removed.

## Frozen protocol

The four cases are, in order:

1. `xlsx_producer_medium_source_full_cell_scan` — the primary 4,096-cell
   producer XLSX corpus;
2. `xlsx_full_cell_scan` — the marker-free XLSX control;
3. `docx_producer_source_selected_paragraph` — the producer DOCX control;
4. `pptx_producer_source_selected_slide` — the producer PPTX control.

The native lane has eight paired blocks per case, 50 measured samples and
three warmups per child. The allocator lane has four paired blocks of the
primary case, 15 measured samples and three warmups per child. Each child is a
fresh sequential process pinned to CPU 12 with a warm filesystem cache. Roles
alternate AB/BA by `(block + case_index) % 2` in the native lane and by block
parity in the allocator lane. This is 64 native plus eight allocator children,
72 children total. The bootstrap seed is 894, with 10,000 draws and retained
indices 249 and 9,749. The run used a shared host with no exclusive-host
guarantee; host activity remains a source of variation for tails and process
descriptors such as RSS.

The primary adoption gate follows the actual 0891 decision: native primary
p50 paired median ratio at most `0.98`, and bootstrap 95% upper bound below
`1.0`. A per-pair ratio above `1.05` is a review flag only; it is not the
adoption gate. Native p50 spread above `1.20` is retained as a host-variation
flag. RSS and allocator counters remain descriptive observations without an
automatic adoption gate.

## Semantic contract

The control corpora and producer sidecars are the sealed 0887 fixtures. The
primary report must also satisfy the 0893 contract on every native and
allocator child:

- `guard_contract` is
  `post-endpoint-per-worksheet-tolerance-v1`;
- measured forced worksheet counts are exactly `[4] * samples` and all
  measured clean flags are true;
- warmup forced worksheet counts are exactly `[4] * warmup` and all warmup
  clean flags are true;
- `forced_worksheet_guard_iterations == warmup + samples`;
- measured `sample_indices`, operation metric indices, and guard vectors are
  aligned after elapsed-order sorting;
- every measured `source_version_stable` value is true;
- timing scope names the process/allocation endpoint snapshots, the
  post-timing per-worksheet guard, aggregate tolerance report, and
  post-guard source-freshness fence.

The verifier keeps the fixed 4,096/3,686/410 cell oracle, archive digest,
marker census, callback digest, producer sidecar, and the exact 0893 static
worksheet-route census. The timed harness does not instrument route selection:
`runtime_route_observed` must remain false and `physical_cold_eligible` must
remain false in both roles. Separate raw-scanner tests supply route eligibility
evidence; a profile report cannot turn that static evidence into a runtime
claim.

## Helper interfaces

`experiment.py --plan PLAN --output-root CAPTURE` launches the frozen 72-row
schedule. It is create-only and fail-fast; a nonzero child or semantic
mismatch stops the schedule without retrying or replacing that child.

`reduce.py --plan PLAN --experiment-root CAPTURE --output REDUCTION.json`
reopens every child through the reviewed 0887 reducer and uses the local
experiment validator, seed 894, and 0894 schema.

`verify.py --plan PLAN --experiment-root CAPTURE --output VERIFICATION.json`
performs an independent no-workload custody replay. It checks every command,
binary identity, output descriptor, interval, schedule row, semantic receipt,
control oracle, and primary 0893 guard vector.

The capture runner loads and validates plan-bound artifact identities before
launching the sequence. A reviewer also called `load_plan` during active
capture, producing extra artifact hashing. Its exact overlap and effect were
not measured; there was no exclusive-host guarantee. The recorded result is
used to reject the candidate, with no precise causal or universal performance
claim.

`freeze.py` additionally requires a create-only binding manifest. Its exact
18 keys are `source_archive`, `source_manifest`, `source_identity`,
`candidate_freeze`, `baseline_build`, `candidate_build`,
`baseline_native_depfile`, `baseline_allocator_depfile`,
`candidate_native_depfile`, `candidate_allocator_depfile`,
`baseline_flags`, `candidate_flags`, `baseline_quality`, `candidate_quality`,
`source_review`, `candidate_review`, `methodology_review`, and `quality_source`.
Every entry is
an object with the exact `kind` and role metadata in `experiment.py`; its
regular-file bytes, mode, and SHA-256 are copied into the plan. Depfiles name
their binary role and build binding, flags name their build, and quality/review
entries name their role or review scope. The verifier reads the linked
receipts and rejects a file-only collection whose build, depfile, flags,
archive, candidate freeze, quality, or review relationship is not proven.

The source contract is exact: schema/status are fixed, revision/base/source
revision are `506aa2eda`, the selected-source list is the ten-entry list in
`source.py`, every four file-count fields equal `10,771`, and the retained
archive descriptor is the exact 89,047,170-byte archive with its frozen
SHA-256. The candidate freeze must be
`litchi.xlsx-selected-metadata-candidate-freeze-0894.v1`, contain the five
reviewed paths, and report a `10,773`-file candidate census (three
replacements and two additions). The plan also seals descriptors for all
seven 0887 report/producer fixtures and the 0893 guard reference, so mutable
historical oracles cannot be swapped after freeze.
It also records exact hashes for the seven packet helpers, including the
transitive 0887 reducer and isolated quality-source preparer; changing any
helper after freeze invalidates the plan. The `quality_source` binding seals
the isolated quality archive, five overlay identities, support-tool/ADR
identities, private lock, and quality-runner copy.

Both baseline build snapshots must independently match the complete
`source.json` 10,771-file census. Candidate `source-before.json` and
`source-after.json` must independently match that census after applying all
five identities from the candidate freeze, for exactly 10,773 files; the
candidate build also binds the exact frozen patch descriptor. Quality and
candidate-review path lists may retain their own traversal order, but each
must contain the unique exact five-path set. Verification rejects symlinked
plan, capture, receipt, binary, and artifact paths before resolving them, and
cross-checks each child receipt's command, working directory, environment,
exit code, and start/finish times.

The plan is created only after the source/build agent has sealed the following
interfaces:

```json
{
  "schema": "litchi.xlsx-selected-metadata-experiment-0894.v1",
  "status": "frozen",
  "base_commit": "506aa2edae6666d9c76ae6827c2f5941a8e47896",
  "guard_contract": "post-endpoint-per-worksheet-tolerance-v1",
  "source": {
    "schema": "litchi.xlsx-source-snapshot-0894.v1",
    "status": "pass",
    "baseline_commit": "506aa2edae6666d9c76ae6827c2f5941a8e47896",
    "timed_harness_guard_contract": "post-endpoint-per-worksheet-tolerance-v1",
    "selected_file_count": 10771,
    "candidate_overlay": {}
  },
  "binaries": {
    "before": {"native": {}, "allocator": {}},
    "after": {"native": {}, "allocator": {}}
  },
  "fixtures": {"xlsx_primary_report": {}, "xlsx_primary_producer": {}, "xlsx_marker_free_report": {}, "docx_report": {}, "docx_producer": {}, "pptx_report": {}, "pptx_producer": {}, "guard_reference": {}},
  "helper_hashes": {"experiment.py": {}, "freeze.py": {}, "verify.py": {}, "reduce.py": {}, "reduce-0887.py": {}, "quality.py": {}, "quality_source.py": {}},
  "bindings": {"source_archive": {"kind": "source_archive", "role": "source"}, "source_manifest": {"kind": "source_manifest", "role": "source"}, "source_identity": {"kind": "source_identity", "role": "source"}, "candidate_freeze": {"kind": "candidate_freeze", "role": "candidate"}, "baseline_build": {"kind": "build_receipt", "role": "baseline"}, "candidate_build": {"kind": "build_receipt", "role": "candidate"}, "baseline_native_depfile": {"kind": "depfile", "role": "baseline", "binary_role": "native", "build_binding": "baseline_build"}, "baseline_allocator_depfile": {"kind": "depfile", "role": "baseline", "binary_role": "allocator", "build_binding": "baseline_build"}, "candidate_native_depfile": {"kind": "depfile", "role": "candidate", "binary_role": "native", "build_binding": "candidate_build"}, "candidate_allocator_depfile": {"kind": "depfile", "role": "candidate", "binary_role": "allocator", "build_binding": "candidate_build"}, "baseline_flags": {"kind": "build_flags", "role": "baseline", "build_binding": "baseline_build"}, "candidate_flags": {"kind": "build_flags", "role": "candidate", "build_binding": "candidate_build"}, "baseline_quality": {"kind": "quality_receipt", "role": "baseline", "review_binding": "source_review"}, "candidate_quality": {"kind": "quality_receipt", "role": "candidate", "build_binding": "candidate_build", "candidate_binding": "candidate_freeze"}, "source_review": {"kind": "review", "review_scope": "source", "source_binding": "source_manifest"}, "candidate_review": {"kind": "review", "review_scope": "candidate", "candidate_binding": "candidate_freeze"}, "methodology_review": {"kind": "review", "review_scope": "methodology", "source_binding": "source_review", "candidate_binding": "candidate_review"}, "quality_source": {"kind": "quality_source", "role": "candidate", "candidate_binding": "candidate_freeze", "archive_binding": "source_archive"}}
}
```

Each binary descriptor must contain an executable path, byte count, mode and
SHA-256. Both build receipts must expose equal flags: Rust 1.95, release,
`opt-level=3`, debug information level 1, fat LTO, one codegen unit, and two
build jobs, plus the exact guard contract. The source receipt must prove a
fresh 506aa2eda selected-source archive and candidate overlay. The packet
cannot be run with a retained pre-0893 executable.

The packet records a scoped measured rejection for the current source, build,
corpus, and protocol; it makes no general Office performance claim. Retained
source overlays, receipts, and raw reports remain available for replay.
