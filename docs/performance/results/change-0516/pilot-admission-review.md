# 0516 pilot admission review

`performance_claim: none`

`performance_admitted: false`

`verdict: provisional reject`

This review covers the completed native before/after pilot for the 0516
emitted-output worksheet parser feed. The pilot is descriptive evidence; it is
not a formal regression study and does not authorize a speedup or a no-op
improvement claim.

## Decision

Reject performance admission for the current snapshot. The protocol's
approximately 5% p50/mean and whole-child RSS gates are exceeded on the
eligible changed-output path, with no independent explanation in the retained
evidence. Do not advance this snapshot to formal ABBA lanes or publish a
performance claim. Correctness and the narrow resource review remain separate
decisions.

The comparison used `metrics.collect()` from
[`metrics.py`](metrics.py), with matched corpus rows, 20 samples, 2 warmups,
CPU 2, one fresh process per capture running all rows and samples, and an
operation clock that excludes setup, expected output, sink reservation,
oracles, and result dropping. Process isolation applies to the capture as a
whole; it does not reset global configuration for each row or sample. The
normal pilot receipts all report successful completion and identical corpus
catalogs. The collector's overall status is currently `partial` because the
formal `r1`/`r2` lanes are intentionally absent. Both allocator repeats are
complete; the current `metrics-summary.json` contains 24 captures and 312 rows
with `issues: []`. The profile capture is still running. Allocator and profile
results remain diagnostic and cannot turn this descriptive pilot into formal
regression confidence.

## Changed-output results

The main pilot has 12 changed-output rows. Every p50 delta is positive, and 10
of 12 exceed the 5% adverse threshold.

| Row | Before p50 (ns) | After p50 (ns) | After minus before |
| --- | ---: | ---: | ---: |
| tiny / one-cell commit | 165,030 | 176,976 | +7.239% |
| tiny / one-percent commit | 301,686 | 324,061 | +7.417% |
| tiny / one-cell commit-save | 205,526 | 218,161 | +6.148% |
| tiny / one-percent commit-save | 369,656 | 393,161 | +6.359% |
| medium / one-cell commit | 1,757,072 | 1,902,973 | +8.304% |
| medium / one-percent commit | 7,034,818 | 7,549,501 | +7.316% |
| medium / one-cell commit-save | 2,229,253 | 2,365,450 | +6.110% |
| medium / one-percent commit-save | 8,856,921 | 9,356,364 | +5.639% |
| dense-wide / one-cell commit | 108,611,714 | 115,571,782 | +6.408% |
| dense-wide / one-percent commit | 221,619,520 | 235,050,794 | +6.061% |
| dense-wide / one-cell commit-save | 155,475,581 | 162,792,664 | +4.706% |
| dense-wide / one-percent commit-save | 313,067,260 | 327,014,709 | +4.455% |

The unchanged public guard separates source-backed no-op scenarios from
changed scenarios. Its six warm changed rows also exceed 5%: the p50 range is
5.821% to 8.639%, with 6 of 6 rows adverse.

| Row | Before p50 (ns) | After p50 (ns) | After minus before |
| --- | ---: | ---: | ---: |
| tiny / warm changed one-cell | 124,371 | 134,541 | +8.177% |
| tiny / warm changed one-percent | 225,381 | 244,851 | +8.639% |
| medium / warm changed one-cell | 1,307,916 | 1,394,436 | +6.615% |
| medium / warm changed one-percent | 2,659,650 | 2,840,582 | +6.803% |
| dense-wide / warm changed one-cell | 80,242,823 | 84,913,644 | +5.821% |
| dense-wide / warm changed one-percent | 165,550,035 | 175,590,692 | +6.065% |

The guard's cold and same-value rows exercise the existing source-identity
protections. Any lower descriptive values in those rows do not offset the
eligible changed-output regressions, and this review makes no no-op performance
claim.

The forced-fallback pilot is near neutral on p50, ranging from -0.443% to
+2.497%. Its rows do not show a measured speedup that could offset the
eligible-path loss.

| Row | Before p50 (ns) | After p50 (ns) | After minus before |
| --- | ---: | ---: | ---: |
| tiny / warm changed one-cell | 488,472 | 497,432 | +1.834% |
| tiny / warm changed one-percent | 941,503 | 953,124 | +1.234% |
| medium / warm changed one-cell | 6,233,564 | 6,314,917 | +1.305% |
| medium / warm changed one-percent | 12,537,658 | 12,482,152 | -0.443% |
| dense-wide / warm changed one-cell | 385,629,540 | 387,964,893 | +0.606% |
| dense-wide / warm changed one-percent | 770,354,356 | 789,591,841 | +2.497% |

## RSS

RSS is GNU `time -v` whole-child resident set size from the normal native
captures; allocator RSS is excluded. The fallback pilot exceeds the 5% RSS
gate even though its elapsed p50 remains within the latency threshold.

| Family / lane | Before RSS (KiB) | After RSS (KiB) | After minus before |
| --- | ---: | ---: | ---: |
| main / pilot | 137,196 | 137,604 | +0.297% |
| guard / pilot | 165,104 | 170,600 | +3.329% |
| fallback / pilot | 246,048 | 261,800 | +6.402% |

The corresponding preflight RSS deltas were +0.506% for main, -0.076% for
guard, and +0.030% for fallback. Those one-sample observations do not explain
the repeated changed-row p50 pattern. Whole-child RSS cannot localize the
allocation owner, so it is reported as a gate result rather than a causal
memory attribution.

## Receipt bindings

The pilot reports are bound to the before source manifest
`2482a2f009ab9213d663a6f8a62607398a72006c13f451a56da8dc78ba80ee5c` and the
after source manifest
`38007ba3313a420bf1f5823fa8aafb2232d9f3f74bc0874ab416e302263b1df6`. The
after candidate is based on `c5abaef129f5ae0000a925b295cf6507b7474e3c` with
candidate patch SHA-256
`9b3cf124904ed5970a376c26a4aabd900c212d34755490f916c60f367c8f9b3f`.

| Pilot | Before receipt SHA-256 | After receipt SHA-256 |
| --- | --- | --- |
| main | [`before/pilot-receipt.json`](before/pilot-receipt.json) `62d034e57743c99b501ac8fd65495b4acea52ff12ef250ee9db68ff68133cdb1` | [`after/pilot-receipt.json`](after/pilot-receipt.json) `33246e3cd66e189fd0d56b74d8be157dbe60c4d64e666cf27ab368cd00e6d635` |
| guard | [`before/guard-pilot-receipt.json`](before/guard-pilot-receipt.json) `8f93182e7e9fb06316d4f7c0ba04a66a3c650c11660d7dffc69375ebe4b8f1b3` | [`after/guard-pilot-receipt.json`](after/guard-pilot-receipt.json) `5c99741dcd6438aa290cb5c2ca2d3ce347e33435671c1c6ca95613aec558e2df` |
| fallback | [`before/fallback-pilot-receipt.json`](before/fallback-pilot-receipt.json) `bbd8ef44bcd474130a3c0473ae6b71e1dbcf4831a1d1ec7134692e452086fe55` | [`after/fallback-pilot-receipt.json`](after/fallback-pilot-receipt.json) `98f1eeaadb356af605737a610213f99b64e27d15fd05a0e64c0a1843a0d530cd` |

Each receipt also binds the executable and its output artifacts:

| Pilot | Before binary / after binary | Before report / after report | Before log / after log |
| --- | --- | --- | --- |
| main | `58ce287eb09dd5fa107be2baf02051b2b89d5146e323866075ee42004efb756f` / `7287f108b2fa8b9810d34896b893c9a0b025e8e96aeab39ac7812b2273919e07` | `197cc8a91cc8f8fc3eb6cb4d6781165b2345cce588f6a4bab5f120d03a9cdc7a` / `13faddaeba18e1bc486cd8b0a29cadd165e7b58288b2abe2eab1a9c9a7bb70c7` | `dfbe7375c01ef0850f0c6dc1e90a675dc232a01b8cbe1febe75f5793d6134912` / `1997214d20b207967f5816d3e7c2fce62332c983210b077b2c447e5cf7f42099` |
| guard | `1077267e9efb1a7be0285472a2753bfa646cf9c2f7d4025571b380598e9d0b71` / `68851618f002458ef4dc66d0841f10700b9d0706dcb114f0582242c7bb003543` | `d920546a40d86fffb60be8b5e4445b4000565bbabc2d706d04c45a1ac6d3b46b` / `4487020d53a9ba54a18208688773b8aae3351c07a0572e093a214173a2f114ef` | `91e9b42b93ca0cd750b5a742a647fb20279748299d401b4ee7176ac4bdc53213` / `9a6a673acc78f7f7aa2fe495ff8b229595102a8a0a1ebd8e9c3c261ea12ed7f8` |
| fallback | `955a8f1dbf8de15d4c3bd64c5aad3c514913d96e9086823cd5529aedaadf191c` / `6f3441ea185bbc85492df656f62c23b6549de3e862b52f94ce1e1dc014737c0f` | `d1184f37c639a034f27dc286b30a91fd1b036a9ad986c744c85123e12c989f90` / `e907efd27c064e1aa4df52cc082d2b8cffafd478148a2a89889d54ee6cbb6409` | `034d5dd61cb632f6be4b2c8cec5449d31d014cd9debf7cc111a86493c5c07a4a` / `87338f2f9a3cd1c8529ff8eff9a64768cb8bfc803778ce27f7030fe6fc5c0c0d` |

The main pilot's corpus catalog artifact is
`2b58a5e6a6c054c0f29b4d95b72381b95c544d6e2b49407f4fa62ae521254ddc` for
both roles. The guard manifest is
`3e4df7e2bd766ca0337acc6f6e28ec0166deea1dc04b8dc566992fc37274e57e`; the
fallback guard manifest is
`b3ea1457eeab4add292ce4dddb913dff23605ddbb73cf4349e1bc8cfb29e4271`.

## Evidence limits and next decision

The pilot has no formal ABBA lanes. Both allocator repeats are complete and
the current `metrics-summary.json` contains 24 captures and 312 rows with
`issues: []`. Allocator and profile results remain diagnostic; the profile is
still running, and these lanes cannot convert the descriptive rejection into
formal regression confidence. The static correctness admission is green, and
the narrow resource review is separately approved, but neither clears the
measured changed-path latency or fallback whole-child RSS gates.

The next admission attempt must first explain or remove the eligible changed
path overhead and the fallback RSS increase, then repeat the same source-bound
pilot. Formal lanes are appropriate only after that pilot passes the retained
latency and memory gates.
