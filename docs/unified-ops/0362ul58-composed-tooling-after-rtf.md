# 0362ul58 — composed tooling after the RTF adapter repair

The current 35-module tooling row passes on clean integration
`1bc3769f43cbe4e2256a7fcc5da7b5530e36f5a0`: **1,689 tests in 107.073 seconds**,
exit zero. The supervising row takes 108.65038967132568 seconds. This is row
48 of the 53-step gate; it does not establish the other rows or native
performance acceptance.

Record 0362ul57 preserves the preceding 1,688-test failure on `319690ffc6`.
The repaired DrawingML preparation adapter keeps the current RTF constructor
spellings in the candidate and restores exactly two historical spellings in
the baseline copy. The historical full-file digest is reproduced exactly;
all DrawingML operation bytes and other non-lock adapter files remain equal.
Existing preparation receipts remain historical and must not be relabeled.

Root reviews the correction, runs the 14 focused protocol tests and verifies
the complete successful row's command log, Cargo trace, source and unchanged
lock identities. Target size remains zero bytes; the Cargo trace is empty.
Every terminal process snapshot has an empty owned-process list. The source
stays clean at the held commit, and the integration hold is released after
verification. No Cargo, native probe or timing workload runs in this row.

External artifacts are retained under
`(deleted-coord-artifact)`. No large log is committed.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `tooling-row-1bc3769f43-20261025T-state/state.json` | 19963 | `27cace6706148d5df0c1889d89fd7c99a0f3232366f57be8e2a123843c34588f` |
| `tooling-row-1bc3769f43-20261025T-state/plan.json` | 24635 | `5d86c0ebdbc12c54b81f8f102c396e9492c1c9533d9ce3a730695bff1a5abb54` |
| `tooling-row-1bc3769f43-20261025T-wrapper.log` | 3998 | `b32419c90df59ff3de6f0b4e88f1fd44df9e7a7319095b9999a7e108258d7bf1` |
| `tooling-row-1bc3769f43-20261025T-source-freeze.txt` | 1224 | `8bc67605660b9991f5166665cc1bc674f6eb972f60824bb3f6db3ad8aaa5c036` |
| `tooling-row-1bc3769f43-20261025T-state/48-tooling-unit-attempt-1792891194503166781-1433696.log` | 635287 | `6cedafcfc5a85d1d04cccc722cc2720923a55a3fccc7a17024fe6dc2edc50b3e` |

The corresponding Cargo-target TSV is empty, SHA-256
`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`.
The state and source-freeze artifacts retain exact commands, runner identity
and the three observed lockfile identities. The earlier failure artifacts
remain unchanged at the paths recorded in 0362ul57.
