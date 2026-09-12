# 0518 DOCX publication profile protocol

This is the source-bound 0518 review for the unchanged baseline and the
matched candidate. The candidate objective is to reuse the caller-retained
OPC source proof and DOCX current snapshot during changed publication while
preserving fresh source authority, limits, security checks, stale-source
behavior, candidate readback, and output guards. The six-arm × two-repeat
candidate profile comparison is complete; full admission still depends on the
native and policy gates recorded below.
OLE2/OOXML are in scope; ODF and iWork remain deferred/outside this review.

## Baseline identity and matrix

The frozen baseline is the unchanged managed DOCX benchmark at revision
`afb62ab7a70859dad4a4b9c8eea91402d1ef4052`, built with the repository Rust
toolchain (`rustc 1.95.0`) and release debug information disabled. The binary
SHA-256 is
`637df10f921f47c7572a25a9900b537ab28a316f85c4ee7010764ff2ecd29255`; the
source-manifest SHA-256 is
`409fa85fd395fff77c7c133e0d98b150ab82fe7efa9efc3a1935a30e9f238e37`.
Baseline profiles are pinned to CPU 2; candidate profiles use their separately
bound candidate binary and source-manifest hashes in the capture receipts and
comparison JSON.

The source-bound native harness is
`crates/litchi-docx/examples/managed_paragraph_batch_perf.rs` with:

```
--paragraphs {128|512} --replacements {1|8|32}
--source {owned|file} --mode {repeated|batch}
--warmups 3 --samples 30 --repeats 2
--artifact-dir <fresh-directory> --output <new-file>
```

The formal native guard remains 24 rows: both paragraph sizes, K=1/8/32,
owned/file sources, and repeated/batch routes, in two reverse-order campaigns.
The bounded Callgrind baseline is the six owned-source arms below, in
`profile-r1` and `profile-r2` (12 profiles total):

| Arm | Reason |
| --- | --- |
| p128 K=1 repeated/batch | flagged small publication case and route control |
| p512 K=1 repeated/batch | larger K=1 control |
| p512 K=32 repeated/batch | established large-selection control |

Each profile uses `--warmups 0 --samples 1 --repeats 1` in a fresh process.
This is a boundary and attribution check, not a stable CPU-performance
estimate or a replacement for the native 30-sample guard.

## Timing and profile boundary

Fixture construction, unmanaged/managed preflight, expected-output creation,
and correctness oracles are outside the native operation clock. The native
`publish_ns` field covers the source-backed publication call and the drop of
its returned `Snapshot`; commit destruction is a later phase. Whole-child GNU
time RSS is one high-water value per child. Hardware `perf stat`, where
available, includes setup, preflight, all samples, verification, and report
handling, so it remains supplementary whole-child evidence.

The Callgrind command is:

```
valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect='*litchi_docx::source_backed::Package::publish_document_commit_to_stream' \
  --zero-before='*managed_paragraph_batch_perf::run_sample' \
  --callgrind-out-file=<profile> <frozen-binary> \
  --paragraphs <N> --replacements <K> --source owned --mode <repeated|batch> \
  --warmups 0 --samples 1 --repeats 1 \
  --artifact-dir <fresh-directory> --output <new-file>
```

Callgrind toggles collection on entry to the selected publication function and
off on exit. The two preflight publications therefore form two complete
on/off intervals; the `run_sample` zero point clears them while collection is
off, and the one measured publication forms the only retained interval. A
single sample per fresh process avoids ambiguity from resetting at each later
sample. The raw proof must show exactly one positive incoming publication edge
from `managed_paragraph_batch_perf::run_sample` with one call, and its cost
must equal the raw `summary` and the selected function's self plus direct-edge
costs.

`analyze_profiles.py` performs that raw check and emits
[`profile-analysis.json`](profile-analysis.json) and
[`profile-after-analysis.json`](profile-after-analysis.json). For every
profile, both the exact-one-call and publication-total-scope checks pass. The
retained annotations are listed in
[`profile-r1/profile-annotations.json`](profile-r1/profile-annotations.json),
[`profile-r2/profile-annotations.json`](profile-r2/profile-annotations.json),
[`profile-after-r1/profile-annotations.json`](profile-after-r1/profile-annotations.json),
and [`profile-after-r2/profile-annotations.json`](profile-after-r2/profile-annotations.json);
each of the 24 profiles has inclusive and exclusive
`callgrind_annotate --auto=no --threshold=100 --tree=both` output.

The publication method ends before the caller drops its returned snapshot, so
selected inclusive Ir is a method-scope diagnostic and cannot be equated with
native `publish_ns`. Callgrind Ir is an instrumented instruction-read count,
not elapsed CPU time, hardware cycles, allocation count, or publication-local
RSS.

## Baseline attribution

The inclusive annotation and raw direct edges agree for all 12 profiles. The
physical publication owner is the direct
`litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream`
edge. The XML validator is
`litchi_opc::xml_splice::validate_source_xml`. On this baseline the direct
DOCX current-snapshot owner is
`litchi_docx::source_backed::Package::main_document_snapshot`; the helper
accepts the current candidate spelling
`litchi_docx::source_backed::Package::main_document_snapshot_with_before` as
well as the intermediate `_with_hint` spelling. A candidate hit may have no
such direct edge when the handoff is inlined.

| Arm | Publication inclusive Ir | Snapshot owner Ir | Topology owner Ir | XML validator Ir |
| --- | ---: | ---: | ---: | ---: |
| p128 K=1 batch | 5,467,227–5,470,253 | 2,952,989–2,955,762 | 2,504,374–2,504,592 | 2,386,179–2,388,327 |
| p128 K=1 repeated | 5,468,080–5,468,556 | 2,954,138–2,955,183 | 2,503,796–2,504,260 | 2,386,475–2,387,027 |
| p512 K=1 batch | 17,746,149–17,747,024 | 11,654,429–11,654,851 | 6,078,594–6,079,869 | 9,454,795–9,455,750 |
| p512 K=1 repeated | 17,743,288–17,743,774 | 11,653,260–11,654,362 | 6,077,416–6,077,878 | 9,455,719–9,456,930 |
| p512 K=32 batch | 17,751,271–17,752,919 | 11,658,823–11,659,806 | 6,078,719–6,081,470 | 9,456,459–9,456,690 |
| p512 K=32 repeated | 17,752,841–17,755,856 | 11,661,135–11,662,253 | 6,078,682–6,080,685 | 9,455,337–9,461,537 |

The snapshot subtree is approximately 54% of p128 publication Ir and 66% of
p512 publication Ir in this baseline. This quantifies the candidate target;
it is not a candidate speedup claim. The topology and validator values record
the already-retained 0517 duplicate-validation optimization and provide the
control boundary for 0518 snapshot reuse. The semantic owner selector accepts
`main_document_snapshot`, `main_document_snapshot_with_hint`, and
`main_document_snapshot_with_before`; the last spelling is the current live
candidate owner.

## Candidate handoff and guard

The completed comparison pairs the same six route names and repeat labels
against baseline, verifies identical CLI settings and source/provider/workload,
and binds each side to its own binary/source-manifest/candidate-plan hashes.
Use `compare_profile_lanes.py` to retain publication inclusive Ir, direct
topology-owner Ir, semantic snapshot-owner Ir, topology inclusive Ir, and XML
validator inclusive Ir. The baseline requires one direct semantic snapshot
owner edge. A candidate hit may inline the source-XML-with-hint handoff into
the snapshot/publication path and therefore have no direct snapshot-owner edge;
the helper reports that as absent instead of treating it as zero instruction
cost. Require raw and inclusive/exclusive annotation identity on every pair
before interpreting a delta.

For each candidate hit profile, verify from the positive raw function costs that
`ensure_source_document_xml` and
`document::transaction::Snapshot::from_source_xml` are absent. This is the
fresh-scan/fresh-snapshot guard; it does not require either helper to appear in
the optimized path. Separately require exactly one positive raw call to
`litchi_opc::xml_splice::validate_source_xml` and retain its positive
annotation row as the final XML validation guard. A fallback or miss campaign
would need its own route label and should not be folded into the hit
comparison.

The result is retained in
[`profile-comparison.json`](profile-comparison.json) and
[`profile-comparison.md`](profile-comparison.md). Across all 12 pairs,
publication Ir falls about 54% for p128/K=1 and 66% for p512/K=1 or K=32;
the topology owner changes by at most 0.05%. The candidate owner is
`main_document_snapshot_with_before`, its fresh DOCX scan and
`Snapshot::from_source_xml` have no positive raw cost, and the final XML
validator has exactly one positive call in every candidate profile. These are
method-scope Ir diagnostics; native elapsed-time and RSS admission remains
separate.

The candidate must still be admitted by the 24-row native matrix, two reversed
campaigns, whole-child RSS guard, correctness/readback oracles, focused OPC and
DOCX tests, and resource/error/security checks. A profile delta can explain a
mechanism, but it cannot replace native elapsed-time/RSS admission or prove
that a reused snapshot is safe on a stale, foreign, byte-mismatched, signed,
encrypted, cancellation, limit, or source-version path.
