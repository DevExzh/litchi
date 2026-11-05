# 0362ul165u — Feature matrix and DOCX qualification gate coverage

The current-7c diagnostic passes all 31 feature-matrix profiles at row 39.
The row completes in 852.96 seconds and leaves approximately 5.01 GB in its
target. Root rehashes the command log and Cargo target trace, verifies clean
source and all four unchanged locks, and checks the resource guards and
post-exit process samples. The diagnostic proceeds to feature-matrix
Clippy. This is evidence for this diagnostic row, not final all-65 acceptance.

The earlier Python row's signed individual no-op assertion is repaired and
merged at `2712b489a3` (165t). Its original failed result remains preserved;
the running diagnostic still uses its immutable `7c043ea168` source.

## DOCX gate integration

The explicit tooling-unit list omitted the newly added DOCX qualification
regression module. Leaf `38629e40b51c5cab37bfe90402210f0edcf08188` registers
`tools.test_docx_metadata_native_qualification` in that existing row and
updates record 165r. Root verifies that the gate still has exactly 53 rows
and includes the module exactly once. The prior independent review of
`bdb6343c3f` remains applicable to the unchanged runner/validator repair;
this successor changes only the gate list and record.

The production crates and performance consumer are byte-identical to the
queued `9c918bd24a` compiler source. The four ordinary/allocator release
builds remain queued behind the diagnostic. No DOCX probe execution or
native qualification is claimed, and the leaf remains unmerged until
compiler validation succeeds.

Root's merge preview of `2712b489a3` and `38629e40b5` is conflict-free, with
virtual tree `c1d7345bbac3742333e13a406259726636865246`. This is neither a
commit nor a selected final candidate.

## Protocol composition return

The v13 packet reconciles the current source identities and 29 required
groups, but does not produce the composed P1/P2 files. Independent review
returns that missing implementation and an incorrect ordering statement.
All source repairs, including DOCX, must merge before selecting final C;
P1 and then the producer-bound P2 follow that source selection.

A v14 successor must carry concrete composed output files and preserve the
previously reviewed transformations and guards. An identity catalogue is
insufficient to prove that composition. Frozen C33/P33 remain unchanged;
actual final-source rebinding and all 29 native groups remain pending.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | SHA-256 |
| --- | --- |
| `remaining20-feature-matrix-root-verification-20261026/receipt.json` | `7060a2d15544906555292357641419feb68643f9d39446ca5c193242bc538836` |
| `docx-metadata-38629-gate-integration-root-review-20261026/receipt.json` | `f550f50309af7a06d05a0aca7785bede043a6c2c390a45e73103a394fc1c507a` |
| `protocol-consumer-composition-C2a-20261026-v13/receipt.json` | `38b97cc3eccf11be50090cbb584f3470d2432c152854883c8c9ea20419603bab` |
| `protocol-consumer-composition-C2a-20261026-v13-independent-review-20261026-fresh/review-receipt.json` | `a9b6de2067bc85bb92cadc0e66263dc7f85c3e99d78915b5bee1f87bb10f3939` |
| `signed-noop-test-retirement-20261026/receipt.json` | `03878322a8315dcc46dad5952e40877cfecc494e27621092c775640baaa4073f` |

Nothing is pushed.
