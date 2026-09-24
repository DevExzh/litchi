# Post-accounting candidate snapshot

This directory retains the exact production source and standalone harness
associated with the `post-*` selector-accounting receipts. It predates the
later ODS conformance fixes now being developed in the shared worktree. The
snapshot is retained for the bounded historical accounting comparison; it is
not presented as the final library source.

The candidate harness source uses the historical fixture captured in those
receipts, including a shorthand second detective range endpoint. The active
harness one directory above has the corrected explicit form
`SheetN.A1:SheetN.A1`; a future final profile must use that corrected harness
for both sides of any comparison. The copied `Cargo.toml` retains the original
repo-relative dependency paths and is intended to be replayed at the active
`harness/` location after materializing the candidate source snapshot.
