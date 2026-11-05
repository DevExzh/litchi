# 0362lb — Share typed Ink graph metadata with the retained owner

- Parent: `62fe6a3f9c`, the 0362la allocation review.
- Status: correctness validated; returned for a remaining no-op peak repair.

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

Matched-v2 source `9aabdcbd2c` preserves all six output files in all four
protected binaries. Four counting legs agree for every route. The
[verified receipt](results/0362lb/evidence.json) records source, dependency,
build and binary provenance. Counts below cover 100 operations; retained
bytes after each route remain zero.

| route | calls, base → candidate | requested bytes, base → candidate | peak bytes, base → candidate |
| --- | --- | --- | --- |
| forward | 307,900 → 304,400 | 36,779,900 → 36,561,600 | 82,323 → 81,923 |
| inverse | 436,300 → 432,300 | 73,803,400 → 73,560,900 | 252,189 → 251,917 |
| no-op | 30,100 → 29,800 | 2,301,600 → 2,273,700 | 3,309 → 4,307 |
| forward + output | 320,600 → 317,100 | 217,352,400 → 217,134,100 | 605,584 → 605,584 |
| inverse + output | 443,100 → 439,100 | 170,786,600 → 170,544,100 | 592,774 → 592,774 |
| no-op + output | 30,200 → 29,900 | 2,660,800 → 2,632,900 | 3,592 → 4,307 |

The duplicate typed-graph allocation is removed, but the no-op patch still
retains a cloned OPC package during fingerprinting. Its higher peak is not
accepted. The next repair must share the existing immutable OPC owner while
preserving isolated changed candidates and source guards. No native timing
run has been attempted for this candidate.
