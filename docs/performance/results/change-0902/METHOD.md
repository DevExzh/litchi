# 0902 premeasurement protocol

The candidate skips a window-search observer only when the existing upfront
condition has already ruled out finding a usable window. Both complete XML
audits remain mandatory, original first; all other branches stay unchanged.
The candidate has not yet demonstrated an end-to-end benefit. Change 0901's
instruction attribution neither establishes branch coverage nor quantifies
observer overhead.

Use fresh paired observations of the retained exact-source 0894 baseline
executables and new candidate executables. Candidate source is the 10,771-file
baseline inventory with exactly two overrides (implementation and parity tests).
Build with Rust 1.95.0, release opt-level 3, fat LTO, one codegen unit, debug 1.
Record source, build, helper, binary, command and environment hashes. Runtime
optional Git/compiler metadata may be null: isolated compile roots have no Git
repository, and workload PATH excludes rustc. External custody is authoritative.

Four fixed cases: XLSX one-percent edit/save on medium and dense-sparse (primary),
XLSX one-edit/save medium, and DOCX semantic one-edit/save medium (controls).
All captures are warm, serial fresh processes pinned to CPU 12. Qualification
uses two samples with zero warmup for each case/role. Main timing uses six paired
blocks, 100 samples and ten warmups per child: 48 children, 4,800 samples.
Alternate baseline/candidate order by block and reverse case order on odd blocks.
Allocator executables are separate: two repeats, three samples and two warmups
per case/role, 16 children and 48 samples. Qualification and allocator timing are
excluded from native benefit estimates. Allocation regions are interpreted only
where the harness actually measures them; RSS is descriptive process residency.

Frozen adoption gate: both primary paired p50 median ratios must be at most 0.98,
with 95% bootstrap upper bound below 1 (10,000 draws, seed 902009). Review each
case's p50/p95/p99 ratio above 1.05 and per-role/case child p50 spread above 1.20.
Keep every flag; do not replace slow rows or weaken thresholds after results.
Require exact semantic/corpus/output/source/sink identities across both roles.
No cold-cache, concurrency, remote-source, broad format, or universal speed claim.

Preserve failed commands and partial artifacts, abort a failing stage without
retries, and version any executed helper before changes. If the frozen benefit
gate fails, restore the two owned production/test files and retain the patch,
tests, quality receipts and measurements as a rejected candidate experiment.

The DOCX control reopens each generated output and checks paragraph text and
commit operation counts inside the unchanged harness; it serializes corpus and
sink identity but no output digest or source counters. Cross-role output digest
and source-counter identity is therefore an XLSX guarantee only. The reducer
must retain this limitation rather than invent DOCX fields. DOCX timing covers
edit/commit/save after package open; it is a shared-code control, not the same
timing boundary as the XLSX primary.
