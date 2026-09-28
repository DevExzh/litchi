# Archived 0820 readers

This directory preserves the planned 0820 offline readers before any capture,
artifact export, or successful quality gate. The archived files are retained
for a future fresh matrix only:

- `analyze.py` — planned 24-selector durability attribution replay;
- `validate.py` — planned successful-capture validator; and
- `reader-notes.md` — the planned-reader contract and interpretation limits.

The original quality gate failed before build/capture on the global allocator
`live_bytes` assertion. The packet's original `plan.json`, `origin.json`, and
`quality-0/` failure evidence remain authoritative for that attempt. There are
zero downstream reports, artifacts, binaries, or captures to analyze.

These readers are intentionally unexecuted and are not evidence for the
repair outcome. A future fresh matrix must copy them back or rebase them onto
its new plan, schemas, source witness, and counts before use; do not run them
against this failed attempt or treat their 216-report contract as a repair
result.
