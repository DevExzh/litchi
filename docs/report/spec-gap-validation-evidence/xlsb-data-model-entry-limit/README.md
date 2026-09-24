XLSB Data Model archive-entry limits
===================================

Data Model source capture derived an archive-member budget above the OPC default total-entry budget, so valid default-limit operations failed with `InvalidReadLimit(ArchiveTotalEntries)`. The adapter now supplies the same checked finite budget for total entries, including ZIP directory entries. OPC admission and overflow checks remain enforced.

The isolated one-line change passed 818 tests across 20 targets (12 ignored), strict all-feature/all-target Clippy, warning-denied rustdoc and formatting. Existing Data Model regressions pass. Independent review is scoped to this limit adapter; the larger Metadata feature and drawing-skipped constructors are separate batches. Exact source and gzip evidence hashes are recorded in `publication.json`.
