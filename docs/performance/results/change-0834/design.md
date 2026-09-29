# 0834 — repair aligned-source filesystem evidence

Base: `41bb4c9670` (0833 failed qualification). Scope is the performance harness,
not production readers/writers. The two retained failures determine this work:
PPTX aligned-source replay classification and OPC cold cross-route byte parity.
The unrelated format review, unified API design and empty matrix file remain
outside the batch. iWork is excluded. All 35 previously read normative inputs
match their 0833 hashes before editing.

First, retain raw PPTX replay request/return observations and child mode while
preserving the failing classifier. Build and run that diagnostic stage against
warm and cold-only invocations. Do not infer which bytes overlapped solely from
the prior error. Freeze/archive the exact diagnostic source and executable;
keep any failures and do not replace their receipts.

Only after those observations prove the cause, implement the smallest bounded
harness correction. The PPTX aligned-source check must prove the exact EOCD
padding transform and identify the exact bounded metadata probe, retaining its
raw overlaps separately from semantic reads. It must still reject any additional
unselected slide/media fetch, unexpected metadata range, source mutation,
incomplete selected payload, or malformed alignment. Normal unaligned behavior
and timed public operations remain unchanged.

For OPC, retain warm route byte parity and prove any permitted cold difference
independently as the private EOCD alignment comment only. Preserve per-route
exact output hashes and semantic/member checks. Do not normalize output, remove
source comments, or relax production preservation. Keep output length/hash,
source identity, comment bytes and logical counters distinct. Any helper must
reject modified payload/framing bytes outside the permitted transform.

Run fresh harness formatting/check/test/Clippy/rustdoc/boundary gates, then fresh
six-case qualification plus the combined OPC save pair which failed in 0833.
Only full report/oracle admission can enable formal capture. The existing plan
of six counterbalanced blocks, 30 samples and three warmups per case/state is
retained for any admitted baseline: 72 reports / 2,160 samples. Native reports
remain route configuration measurements, not a before/after optimization claim.
Record OPC package-drop asymmetry and PPTX untimed replay counter scope.

ADRs 0003/0005/0006/0008 require unchanged publication, bounded evidence,
preservation and admission. 0010/0011/0024 retain package ownership. Harness
alignment helpers parse only their known generated ZIP framing, never becoming
production package grammar or ordinary CRUD APIs. No new ADR is needed.
