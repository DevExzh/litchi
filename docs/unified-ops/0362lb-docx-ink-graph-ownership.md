# 0362lb — Share typed Ink graph metadata with the retained owner

- Parent: `62fe6a3f9c`, the 0362la allocation review.
- Status: correctness validated; fresh allocation and native evidence pending.

The private Ink graph binding no longer copies content-type strings and every
relationship field already owned by its immutable OPC package. It retains the
exact content-type token hash, per-owner relationship token hashes and member
presence, and sorted part names. Replay compares the candidate's typed content
types and relationships directly with the retained package. Exact stored part
spelling is checked even though OPC lookup permits case-insensitive fallback.
Part payload comparison remains unchanged.

Candidate capture still admits the entire graph before reporting a mismatch.
It preserves all graph/token limits and validates internal relationship target
names. Lexical-only changes, empty relationship-member insertion, unrelated
graph or payload changes, dirty state, signatures and retained owner policy
remain covered by the existing Ink integration suite. No public API changes
in this repair, and no default-policy shortcut or global cache is introduced.

All 57 Ink integration tests, 1,489 library tests and 170 doctests pass (31
existing doctests ignored). All-target/all-feature Clippy with warnings denied,
formatting and both fresh DOCX glossary profiles pass. The
[validation receipt](results/0362lb/validation.json) records exact commands,
logs, source hashes and dependency lock. Fresh matched binaries will use the
unchanged six-route probe and builder from 0362l against the original baseline
`d0406be017`. The failed 0362la comparison remains recorded separately.
