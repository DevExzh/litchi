# Repaired harness review

The production readers, writers, and operation timers are unchanged. Only
`tools/perf-baseline/src/filesystem.rs`, its new `aligned_zip.rs` helper, and
the harness README are in the source allowlist. The diagnostic source is
archived separately from the repaired source.

PPTX review found no correctness blocker. The constructor/query boundary is
captured directly after `SourceBackedPresentation::from_read_at`. Rebuilding
all phase counters and coverage from chronological returned ranges must match
the existing aggregate snapshot. Query coverage comes from its own phase,
not from subtracting merged constructor coverage. Only verified aligned modes
can invoke the tail allowance. The proof binds the pinned ordinary corpus,
aligned source hash/size, page size, EOCD position, original comment, and zero
suffix. One exact full tail request is required; additional payload overlap
is rejected. Raw overlaps remain visible in the report.

The focused PPTX tests use synthetic ranges. They cover accepted metadata
overlap, duplicate/missing/short tails, query-phase tail misuse, additional
payload reads, and denial of an allowance for an ordinary source. They do not
replace full pinned-corpus qualification or add duplicate/missing-member
fixtures. Existing payload-range discovery still rejects missing or duplicate
slide slots before classification.

The OPC helper proves each route independently against the aligned archive.
It compares fixed EOCD bytes and unchanged member local/central bytes, verifies
the route's exact comment policy, and retains raw output hash and length.
An untimed canonical digest removes only the already-proven comment difference
for cross-route comparison; it does not modify published output. Warm and
advisory-cold comparisons continue using raw output hashes. Every cold sample
must match its route's raw artifact hash and length. The normal semantic
verifier remains active. The first qualified attempt exposed descriptor-bearing
members in the generated OPC archive. The final helper relies on
`ZipArchive::get_entry` to validate descriptor grammar, then includes bytes
through the next physical local header (or the central-directory boundary)
in each member's raw comparison. Duplicate local starts and payloads crossing
the derived boundary are rejected. Descriptor acceptance, span inclusion, and
mutation tests cover this correction; no second descriptor parser was added.

Root review caught and corrected two issues before the repaired freeze: an
early helper compared only bytes before the EOCD instead of including its
fixed fields, and a synthetic query-tail test used nonoverlapping ranges.
The helper's additional unique terminal-EOCD check is deliberately limited to
the generated ZIP32 fixtures; it is not production ZIP grammar.

Independent report admission, actual verified-cold qualification, quality
gates, and final custody checks remain separate from this source review.
