# Change 0246: optional DOCX semantic resource case

Date: 2026-08-20

Status: tooling and schema evidence only; no profile run or performance claim

## Scope

`tools/perf_resource_profile.py compare-docx-semantic` runs a retained,
process-isolated A1/B1/B2/A2 resource ABBA using explicit control and candidate
binaries. Its default fixed case tuple remains
`docx_semantic_open,docx_semantic_full_text`. The opt-in
`--include-one-paragraph-text` (or
`--include-docx-one-paragraph-text`) flag adds
`docx_semantic_one_paragraph_text` to that tuple for all four legs.

Each selected case must appear exactly once and in the fixed order. Every leg
still requires the same deterministic large DOCX corpus manifest, including
archive and target payload SHA-256 identities, and the same clean harness,
revision, tool, and environment identities. The published report retains the
canonical harness identity even though raw harness JSON is not copied into the
compact report.

## Interpretation boundary

`/usr/bin/time -v` and optional heaptrack observations are whole-process
resource evidence. The per-case harness elapsed summaries are instrumented
resource observations only; `latency_evidence.status` remains `not_measured`.
The report makes no physical-I/O, cache, source-byte, decompressed-byte,
recompressed-byte, or memory-copy claim.

## Verification

Focused Python unit tests cover parser selection, exact optional case and
corpus identity, instrumented metric labeling, time/heaptrack leg wiring, and
published latency separation. No Cargo command or profile run was performed
for this change.
