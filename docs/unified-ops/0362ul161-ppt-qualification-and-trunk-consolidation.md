# 0362ul161 — PPT repair verified and integration consolidated onto trunk

The PPT repair chain through 158h is merged at
`ac36f3099c245920b964b990a3bfa090464955ae`. The coordinator fast-forwards
`feat/unified-ops-api` to that commit, bringing the accumulated integration
work onto the requested trunk. This is integration, not final acceptance:
the corrected full gate, final census, fidelity and performance remain open.

PPT now intersects its source limits with OfficeArt's safety ceiling, retains
the caller's limits through its projections and image admission, and reports
typed limit failures correctly through Ops. Checked standalone picture-bullet
construction and content equality preserve the previous capabilities. The
hyperlink-add case that failed C1 now passes without changing the CRUD oracle.

## Compiled verification

The approved capped runner executes serially with locked, offline Cargo,
one build job and one test thread. Root verifies complete logs and target
traces, unchanged source and three locks, resource guards, and empty owned
processes after each terminal run.

| Source | Result |
| --- | --- |
| 158e `0d4057395f` | Compilation stops at eight strict lints; no tests run. |
| 158f `40eac77e8d` | PPT library: 1,249 pass, one existing ignored fixture test. Ordinary library compilation then catches a test-only constant. |
| 158g `34fc969802` | All nine OfficeArt unit tests pass. New integration assertions fail type inference; no integration test runs. |
| 158h `3b4fd734ea` | Three hyperlink tests, two picture-bullet tests, the typed Ops error test and all-format CRUD pass. Zero unreached claimed cells. |

158f removes stale private default-limit code; 158g moves the remaining
fixture constant into tests. 158h uses an explicit nonempty BLIP payload,
making byte preservation observable while correcting its assertion types.
All changed Rust files pass rustfmt. Failed packets remain intact.

The final focused packet is external at
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-158h-focused-20261026/`.
Its `root-terminal-verification.json` is 2,423 bytes, SHA-256
`acc52ccc6d295ef29ba45b202758d3f91e4f3ea3fe00305f787155a51a0d9922`.
The plan SHA-256 is
`eac241c892ff0c7dac80350dca9a77d89c13b075b53304a5fbdd253e8e7c3057`.
It records the exact commands and retained per-row logs. The earlier full
library and OfficeArt results remain bound to their stated commits rather
than being relabeled as final-candidate results.

## Branch and checkout cleanup

The refreshed census accounts for all 75 pre-cleanup heads: five active
unified controls, the two protected main/office refs and 68 independent
office refs. The merged PPT worktree and exact ref are removed after its
ignored lock is preserved. Its completed test cache is also retired.
There are now 61 registered worktrees and 74 heads; cumulative removals are
277 worktrees and 439 refs. The four retained unified refs are already in
trunk and remain as the trunk, integration, native protocol and gate controls.
Required baseline checkouts remain until their qualification closes.

The checkout had 320,315,392 allocated bytes and its test cache had
3,335,327,744 allocated bytes. These are allocation ledgers, not measured
free-space gains. External evidence under `.coord/cleanup-ppt-158-20261026/`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `events.jsonl` | 1,199 | `ebb7488f0e3019a7f5d2e7f48bed8e319b2cce2f7e9dbb89517fb1a2b42d8fd4` |
| `target-retirement.json` | 467 | `66fd441a12b569e95e59527872f89d814b559c1973d396565bc3eb865142fbd9` |

Before the trunk merge, root preserves the original 15-byte untracked
`constructor-probe.rs` with its metadata. Git supplies the distinct tracked
probe. The ignored root Cargo lock is also preserved before migration to the
qualified candidate lock; both tracked tool locks arrive through the merge.
All three root locks then match the integration worktree. The external
`.coord/root-probe-before-c2-integration-20261026/receipt.json` is 2,409 bytes,
SHA-256 `46e0d85a501436b399e4897a491841ad7f7a43a74121495dd18afce4013e29fc`.
The receipt discloses the initial postcheck's incorrect unchanged-lock
assumption. No source or original local content is lost.

## Remaining qualification

Raw census successor v12a passes independent source review and 18 tests.
It pins the complete 68-package, 71-glossary/13-package-only matrix, rejects
stale profile IDs anywhere, seals the complete driver chain, and preserves
the 88 capture steps plus separate strict finalization. Its external
`raw-census-f5-v12a-matrix-compat-20261026/independent-review-receipt-v2.json`
has SHA-256
`2f50b5b08ffa2e232ee032e77a76fdf47e213cbd0e6fd21a23810289ecb55572`.
This is source preparation; no final profiles have been captured.

The next step freezes a corrected C2 and runs the fresh full65 gate, including
isolated workspace rustdoc, before final native qualification. C1 candidate
bindings remain historical. Native policies 72/74/75, notes and scanner126,
the final census and OfficeCLI postcheck still require corrected-source
results. Memo 22 remains the only deferral. Nothing is pushed.
