# 0362ul66m — Numbers-wire atomic composition proposal

Status: composed on the current integration tip for coordinator review.

## Source and composition

The independently validated leaf is
`67613f5dd0f333cf0b76c3736a980cdc2cc7442d` on
`uo2/numbers-wire-raw-0362ul66`.  Its complete 0362ul66 lane is the ordered
range beginning at `63ab8b4ceae213abe916500017016e8b1fd4eb41` and ending at
that tip.  It covers the raw `Limits`/typed-error contract, every Numbers,
Pages, Keynote and iWork consumer, fuzz targets, tests and fixtures, the
glossary/profile entries, the scoped F5 plan and all compile/test repairs,
including the final tile-output preflight repair in `0362ul66l`.

The range was cherry-picked, in order, onto the current integration tip
`66364d4fb0dc98ab04072f248cd6008f5974346f`.  The composed branch before
this record is `25cda74b33` (`uo2/numbers-wire-raw-0362ul66-compose-current`).
No unrelated commit was cherry-picked from the leaf.

## Source comparison and overlaps

The 96 paths changed by the leaf were compared by blob ID between the
validated leaf and the composed tree:

```text
94 paths identical
0 paths missing
2 shared additive paths differ
```

The two expected shared paths are `docs/API_GLOSSARY.md` and
`tools/api_glossary.json`.  Integration already carries additive rows and
sections for other completed raw-policy lanes.  The composition retains those
entries and retains the Numbers-wire `numbers-wire` profile, required
signatures, raw-boundary table row and `0362ul66` section; no existing
integration row was replaced or deleted.  All 94 production, consumer,
fixture, fuzz, test and lane-record paths are byte-identical to the validated
leaf.

## Validation provenance

The approved scoped F5 is recorded at
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul66-f5-actual-67613f5dd0` and
passed all 16 rows on the validated leaf.  Its source was clean and used the
approved lock digests:

```text
Cargo.lock                         4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
tools/perf-baseline/Cargo.lock     eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c
tools/native-resave/Cargo.lock     51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378
```

Receipt digests are retained in the external state directory: `plan.json`
`6d0379322a2e5614363da3da338f6834ec13cb3335182db7f6a586ccd23235db`,
`rows.json` `a8f2da03c69b62eaac82e41423bd376e0d9e2f86a370b9b1244b9a6af1dfb488`,
and `result.json`
`66749085fa1eb3c604d7d025cb9fa003c96fba318a33b3ea2385fdb78396c5cc`.
The composition changed only the two glossary blobs in addition to the
validated lane content; the coordinator should rerun the source glossary
profiles after merging the composed tree.
