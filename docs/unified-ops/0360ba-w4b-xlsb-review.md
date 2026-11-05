# 0360ba — W4b XLSB review corrections

- Status: prepared on `uo2/w4b-xlsb` for coordinator review after record
  [0360b](0360b-w4b-xlsb-contract.md); the final merge commit is assigned by
  the coordinator.
- Date: 2026-10-21.
- Input: the coordinator review of 0360b, the owner decisions in [0351]
  (0351-api-contract.md) and [0355a](0355a-w1-docx-gaps.md) §1.
- Scope: the XLSB raw and record APIs, sparkline limits, source-backed
  metadata, the part patch engine, and edit-path refusal classification.

## 1. Corrections

The raw constructors now require their limit argument:

```text
raw::Encoder::new(writer, limits)
raw::Records::new(input, limits)
raw::Cursor::new(data, context, limits)
```

`bounded` remains the fallible validating convenience constructor for each
codec. There is no public raw `new` overload with a default limit and no
public `*_with_limits` alias. The same rule applies to
`sparkline::{Formula, Group, Groups}::new` and `Formula::name`: callers pass
the required `sparkline::Limits` directly. All XLSB and facade consumers,
tests, examples and harnesses use the explicit form.

Source-backed metadata now has the complete family shape. A
`SourceBackedWorkbook::edit_metadata` commit returns the deferred core
`Commit`, and `SourceBackedWorkbook::apply_metadata` replays a patch against
the package content base before reopening the deferred owner. The replay
composes source revisions and preserves the source freshness and signature
checks. `metadata_source::source_metadata_patches_replay_against_the_content_base`
covers two successive patches and both inverse patches.

Raw OPC relationship publication retains lexical state. A part patch records
the relationship entries, the exact retained relationship-member bytes and
whether the member existed. Capture compares both the semantic graph and the
retained bytes, so a whitespace, attribute-order or equivalent lexical edit
is a changed patch and is replayed and inverted byte-for-byte. Manifest
replay still covers only the touched owners. `LazyManifests` equality now
compares the actual captured manifest views and never reports distinct lazy
captures equal merely because they have not yet been materialized.
The OPC layer exposes retained XML as a shared byte buffer when its semantic
binding is still valid, so ordinary changed-part capture does not reparse
relationship XML; canonical capture is reserved for an authored or
semantically changed relationship graph. Every source relationship and
content-type fallback used during capture and restore receives the package's
explicit `read_limits()`, so a source admitted under a larger limit remains
editable through replay.

Content-types-only lexical edits are retained as manifest patches, so they
are no longer mistaken for an empty raw OPC edit. Replay preflights an
explicit replacement against its recorded publication and source allocation
before applying typed deltas, then restores the after image against the
post-delta publication. This keeps mixed lexical and structural edits
replayable on an equal package, rejects a divergent manifest instead of
silently dropping the raw bytes, and lets an authored structural fallback
compose with unrelated additions. Source allocation identity distinguishes
an explicit replacement from a structural edit that temporarily invalidates
an older source binding; both directions of an explicit replacement retain
the conflict guard. The patch stores compact shared manifest byte images and
changed-owner relationship images, so retaining patches does not retain
complete before/after package graphs. This keeps structural composition and
lexical fidelity separate.

The regression test
`theme::raw_opc_lexical_relationship_edit_is_reversible` proves that a
relationship-only lexical edit changes the package, replays to the same
publication bytes and inverts to the original bytes.

## 2. Edit errors

Authoring and edit routes that previously exposed free-text
`UnsupportedFeature` or `Unrecognized` errors now return
`Refused(Refusal)`, with a static reason and the appropriate
`RefusalKind`. This includes opaque calculation-chain, data-model, XML Maps,
volatile-dependency, scenario, slicer, sparkline, writer-property,
conditional-formatting, comments, threaded-comments and structural
sheet-edit refusals. A cell-watch source mismatch is a `PatchConflict`.
Read-only parsing and source projection paths retain their format or
unsupported-feature variants so malformed input remains a parse diagnostic;
the edit boundary classifies a well-formed request that cannot be preserved.

The typed refusals preserve the existing atomicity guarantees: failed edits
do not publish a candidate, identity operations remain exact no-ops, and
all refusal reasons are content-free. The corresponding XLSB tests now
assert `Refused` for structural, formula-context, opaque-owner and authoring
cases.

## 3. Validation

Focused validation on this review tree:

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo check -p litchi-xlsb --all-features --all-targets` | 0 |
| `cargo test -p litchi-xlsb --all-features --test sheet_remove` | 28 passed |
| `cargo test -p litchi-xlsb --all-features --test sheet_order` | 15 passed |
| `cargo test -p litchi-xlsb --all-features --test sheet_order_dependencies` | 6 passed |
| `cargo test -p litchi-xlsb --all-features --test sheet_insert` | 13 passed |
| `cargo test -p litchi-xlsb --all-features --test cell_format_formula` | 6 passed |
| `cargo test -p litchi-xlsb --all-features --lib sparkline` | 41 passed |
| `cargo test -p litchi-xlsb --all-features comments` | 18 passed |
| `cargo test -p litchi-xlsb --all-features conditional` | 58 passed |
| `cargo test -p litchi-xlsb --all-features --lib opc` | 13 passed |
| `cargo test -p litchi-xlsb --all-features --lib` | 741 unit tests passed |
| `cargo test -p litchi-opc --all-features --lib` | 694 unit tests passed |
| raw OPC lexical relationship and content-types regressions, including mixed and invalid-source inverse replay | passed |
| source-backed metadata replay regression | passed |

The complete library runs pass after the consumer expectation and
edit-boundary corrections: 694 OPC unit tests and 738 XLSB unit tests. The
focused comments, conditional-formatting, relationship and lexical replay
filters also pass. The broader integration and doctest matrix remains the
coordinator's final gate.
ABBA measurements for this corrected tree remain a coordinator measurement
step; the host was quiesced for the coordinator's isolated probes. The final
paired route measurement is recorded in section 5.

## 4. Compact manifest replay measurement

The compact manifest implementation and the stale-source replay regression
are committed in `122bb3921b1caee1362c36d52593f9cca1dc0a38`. The expanded
probe was built in an independent release target from that source and pinned
to CPU 29. Its binary is 7,196,416 bytes with SHA-256
`0458b2b71cf07aa29c13000238bf677d99849ff2e312890d9c674fc3758326f5`.

The final 31-sample run is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/expanded-run-compact-final-1792580640/head.json`
(605 bytes, SHA-256
`8649037883031697756c8d0ee198e790fce6d744eeb8f394ce9aff8f4a1d9c3d`),
from:

```text
taskset -c 29 /home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/tgt-expanded-after-v2/release/w4b-expanded-probe 31
```

| Case | Median |
|---|---:|
| structural/commit | 11.2142 µs |
| structural/commit_save | 157.5488 µs |
| structural/forward_replay | 9.5820 µs |
| structural/inverse_replay | 9.0640 µs |
| history/retained128_build | 4714.8330 µs |

The complete expanded case list is exactly those five cases. The retained
128-patch heaptrack capture is
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/expanded-run-compact-final-1792580640/retained128.heaptrack.zst`
(106,856 bytes, SHA-256
`c5e9561984b5c52a9b281c0a4bb5a4c816cae6bb3b0d5200299ff2125a5f798a`). It
reports 412,681 allocation calls, 2,541 temporary allocations, one leaked
allocation (544 bytes), 508.23 KiB peak heap and 9.23 MiB peak RSS including
heaptrack overhead.

## 5. Paired final route ABBA

The final compact binaries were measured in a quiet `before -> after -> after
-> before` sequence, with 31 samples per route and CPU 29 pinned. The run is
outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/final-run-compact-abba-1792581200`.
Its command was equivalent to:

```text
taskset -c 29 <binary> 31
taskset -c 29 <binary> --case all --backend owned --fixture test-data/poi/test-data/spreadsheet/testVarious.xlsb --warmup 3 --samples 31 --json <output>
```

The six binaries were recorded in `binary-sha256.txt` (1,025 bytes, SHA-256
`7480e8d203fc3dd952cf23c24afc3ffdcb5d78fbdac42f17652b8e96f41612ba`):

| Route | Before SHA-256 | After SHA-256 |
|---|---|---|
| XLSB probe | `b8398575a088f69c48fdc3b3d10eaa40ad806ed1ac3fabdb6386697256e4fcd6` | `7dcbfd357fb320f31087355a30ae1a389bfe00e7bc277b6543fce841be8a5672` |
| sheet dispatch | `8fa6a5ff09ab634c3ec2f69b4b30e673a1c310088770e87fd399b22e0f2c2ed6` | `b8b3559c78c44346f5dad33efb04ac0249459860dfe9a27d53a98c2d77e796a7` |
| CRUD | `de3021bcef474cb59bb2f21353597c0d86f95546c98e1c407ed6a16cd8127d27` | `90e9e2679a546b91bb02977e777ee14d0ab4268ef6d84db26c3f57eb06ac5a38` |

The probe medians are in microseconds and compare the two before legs with
the two after legs by their median:

| Case | Before | After | After / before |
|---|---:|---:|---:|
| `eager/batch1k` | 2025.975 | 1856.280 | 0.91624 |
| `eager/chain` | 2929.381 | 2692.725 | 0.91921 |
| `eager/comment` | 45.367 | 45.914 | 1.01206 |
| `eager/format` | 1194.500 | 1050.576 | 0.87951 |
| `eager/hyperlink` | 40.103 | 39.393 | 0.98230 |
| `eager/merge` | 958.481 | 924.361 | 0.96440 |
| `eager/metadata` | 35.119 | 34.546 | 0.98368 |
| `eager/name` | 1757.332 | 1610.921 | 0.91669 |
| `eager/open_get` | 190.040 | 176.411 | 0.92828 |
| `eager/set` | 411.674 | 358.798 | 0.87156 |
| `eager/set_save` | 403.623 | 353.960 | 0.87696 |
| `leaf/batch1k` | 1315.032 | 1139.751 | 0.86671 |
| `leaf/open` | 72.020 | 70.591 | 0.98016 |
| `leaf/set` | 435.986 | 355.368 | 0.81509 |
| `leaf/set_save` | 431.081 | 351.242 | 0.81479 |
| `source/batch1k` | 2584.668 | 2416.957 | 0.93511 |
| `source/format` | 1686.392 | 1544.974 | 0.91614 |
| `source/metadata` | 18.701 | 18.754 | 1.00284 |
| `source/open_get` | 157.404 | 149.823 | 0.95184 |
| `source/set` | 758.706 | 709.872 | 0.93564 |
| `source/set_save` | 752.449 | 709.004 | 0.94226 |

The XLSB sheet-dispatch medians are likewise in microseconds:

| Case | Before | After | After / before |
|---|---:|---:|---:|
| `cell/xlsb` | 770.858 | 799.664 | 1.03737 |
| `cell_value/xlsb` | 791.443 | 798.554 | 1.00898 |
| `cells/xlsb` | 290.556 | 302.817 | 1.04220 |
| `evaluate/xlsb` | 1733.563 | 1722.008 | 0.99333 |
| `recalculate/xlsb` | 5609.738 | 5384.676 | 0.95988 |
| `recalculate_heavy/xlsb` | 6106.540 | 6134.770 | 1.00462 |

All CRUD cases improved; their medians are shown in microseconds:

| Case | Before | After | After / before |
|---|---:|---:|---:|
| `edit_ceil_one_percent_existing_cells_save` | 2483.992 | 2142.166 | 0.86239 |
| `edit_one_existing_scalar_save` | 2483.033 | 2147.601 | 0.86491 |
| `full_stored_cell_scan` | 521.953 | 513.507 | 0.98382 |
| `full_text` | 164.810 | 162.616 | 0.98668 |
| `noop_transaction_commit_save` | 926.415 | 568.953 | 0.61414 |
| `open_identify` | 976.745 | 637.083 | 0.65225 |
| `selected_worksheet_cell` | 525.112 | 513.598 | 0.97807 |
| `worksheet_catalog` | 516.263 | 509.073 | 0.98607 |

The full sheet-dispatch matrix is present in the external `sheet-*.txt`
captures. Its unrelated tiny `cells/xls` route measured 1.090x in this run;
the same route measured 1.075x, 1.079x and 1.126x in the three preceding
ABBA runs, while the XLSB routes above remain within the expected variance.
The paired reopened structural/history comparison is recorded in section 6.

## 6. Reopened structural replay and attribution

The earlier expanded candidate-only sample was replaced with a paired probe
that uses independently reopened equal-content recipients for both forward and
inverse replay. The preflight sorts the OPC inventory by part name, compares
each retained part's exact bytes, and then compares the serialized package
bytes. Both binaries report the same preflight:

```text
preflight exact: base_parts=7 after_parts=8 base_bytes=5155 after_bytes=5320
```

The probe sources are outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/expanded-probe/src/main.rs`
(SHA-256
`65d0629300694e7bab1bbc15b053223e58f77db2832bf75523ac3ddf5b7615b4`) and
`.../expanded-probe-before-v2/src/main.rs` (SHA-256
`74f2271ca386e08515e50f5b0d87e9b4d6e514e7284084179114d2571e794a44`). The
baseline uses the old `into_opc`/`from_opc` operation and keeps each published
package snapshot in its history. The candidate uses the new edit/commit and
keeps each `Patch`; this difference is explicit in the history comparison.
The release binaries were built in separate absolute target directories:

| Binary | Size | SHA-256 |
|---|---:|---|
| `tgt-expanded-before-v4/release/w4b-expanded-probe-before-v2` | 7,632,296 bytes | `c38df593f0083aa3c0379d67363796adeb4f97076976393b8f6dddd61e5bce01` |
| `tgt-expanded-after-v4/release/w4b-expanded-probe` | 7,451,432 bytes | `843034d24a33f968d1d5a6698864435d09e0516e576e4750ec75c8a0dc907fdc` |

The four legs ran `before -> after -> after -> before`, pinned to CPU 29,
with 31 samples per timed case. JSON captures are in
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/expanded-run-reopened-abba-1792582900`; their individual SHA-256 values are in
`timing.sha256` (702 bytes). Medians are microseconds:

| Case | Before | After | After / before |
|---|---:|---:|---:|
| `structural/commit` | 10.6060 | 14.4891 | 1.366123 |
| `structural/commit_save` | 44.8822 | 50.2313 | 1.119181 |
| `structural/forward_replay` | 73.1343 | 81.6354 | 1.116240 |
| `structural/inverse_replay` | 75.6574 | 81.5505 | 1.077892 |
| `history/retained128_build` | 4607.5885 | 4700.6240 | 1.020192 |

The retained-history heap runs use direct `heaptrack` invocation and the same
ABBA order. The baseline snapshot history reports 594,920 allocation calls,
1,990 temporary allocations, 4.40 MiB peak heap and 15.25/15.12 MiB peak RSS
on its two legs. The candidate patch history reports 417,161/417,148
allocation calls, 3,190/3,177 temporary allocations, 781.98 KiB peak heap and
10.19/10.17 MiB peak RSS. Both have 544 bytes leaked. These are paired
retention measurements with the old snapshot and new patch history semantics,
so the lower candidate live history is evidence for the compact patch model,
not a claim that the two retained object graphs are identical. The compressed
captures are 159,180–176,477 bytes; `heap-direct.sha256` records their hashes.

The full sheet-dispatch route was profiled separately for `cells/xls` using
the existing before and compact-after binaries, with `SHEET_DISPATCH_ONLY=cells/xls`,
1,000 rows and 31 samples. `perf stat` could not open any event on this host:
`/proc/sys/kernel/perf_event_paranoid` is `4`, and the kernel reports no
`CAP_PERFMON` access. To obtain a deterministic instruction attribution, the
same before/after ABBA was run under Valgrind Callgrind with only `Ir` enabled.
The before instruction counts were 1,287,323,990 and 1,286,565,947; the after
counts were 1,278,277,885 and 1,276,960,259. Their median ratio is
`1,277,619,072 / 1,286,944,968.5 = 0.992753`, so the repeated wall-time
increase of the tiny `cells/xls` route is not an instruction-count increase.
The Callgrind captures are 643,027–653,510 bytes; `callgrind.sha256` records
their hashes. Callgrind emitted `brk segment overflow` warnings while still
producing complete `Ir` summaries; no cycle count is claimed because hardware
cycle counters were unavailable.

## 7. Memoized content-types coverage

The OPC package now memoizes whether its admitted `[Content_Types].xml`
binding still covers the current part graph. The state is clone-local, so a
staged edit cannot invalidate or authorize a sibling package's memo. Mutable
part access, same-name replacement, part insertion/removal/materialization,
source-token restoration, relationship replacement and package/part
relationship access all reset the state before exposing the changed graph;
sign, resign and unsign follow the same rule through their signature-graph
paths. Payload-only replacement leaves the content-types state intact because
it cannot change the mapping.

The cache repair is covered by
`package::tests::content_types_source_state_is_clone_local_and_tracks_mutation_routes`,
which checks clone isolation, borrowed payload access, same-URI retyping,
part insertion/removal, root and part relationship changes, and unsigning.
`package::tests::content_types_source_token_restores_removed_override_after_reopen`
checks that source-token restoration resets and recomputes the memo after an
independent reopen. Existing relationship tests cover absent versus explicit
empty members and reopened lexical restoration; the XLSB lexical root test
also replays and inverts a reopened raw relationship edit byte-for-byte.

## 8. Dependency-closure projection publication

Raw OPC commits, staged edit commits, and `apply_opc` now use the same
`PartPatch::projection_unchanged` proof before publishing a candidate. Root
relationship edits always take the existing full-reparse path. For part
deltas, the proof lowercases the changed part names and walks the internal
relationship closure in the recipient's before graph. With no root delta,
every relationship owner that can differ after apply is already in the part
delta set, so an unchanged reachable owner cannot acquire a new edge to a
changed part; the before-graph proof also covers a recipient's unrelated edge
to a part that was inert in the recorded base. External relationships are
outside the workbook projection; malformed or dangling internal targets fail
the proof conservatively and retain the full-reparse path. A manifest-only or
dependency-disjoint part patch still runs the normal patch, manifest,
signature, source-limit, and lightweight workbook checks, then publishes an
unparsed package so the workbook projection is not decoded twice. Any
closure-affecting change continues through
`reparse_candidate_retaining_projection`, preserving the validating open's
projection and tolerance policy.

The new `opc` tests cover an unreferenced part through forward replay,
independent reopen, inverse replay, and independent inverse reopen. A second
test adds a transitively referenced part and verifies that the raw commit,
forward replay, inverse replay, and both reopened routes remain eagerly
validated. These tests compare a name-sorted semantic OPC inventory and exact
part bytes, content-types publication bytes, and retained relationship bytes;
they intentionally do not treat ZIP member ordering or writer metadata as
part-content differences.

Focused validation passed 13 `opc` tests and the full XLSB library suite (741
tests). `cargo fmt --all -- --check` and `git diff --check` also pass. The
first paired performance measurement of the dependency-disjoint publication
path is recorded below; its inverse leg exposed a redundant post-apply closure
walk, so that candidate is explicitly rejected and the repair is recorded in
the next section.

The first paired measurement after this change is now available. The same
before/after probe sources were rebuilt into distinct absolute target
directories from commit `a6a55b09f4` and its recorded baseline. The baseline
binary is
`tgt-expanded-before-a6a55b09f4/release/w4b-expanded-probe-before-v2`, size
7,643,696 bytes, SHA-256
`c38df593f0083aa3c0379d67363796adeb4f97076976393b8f6dddd61e5bce01`; the
candidate is
`tgt-expanded-after-a6a55b09f4/release/w4b-expanded-probe`, size 7,482,512
bytes, SHA-256
`59b03e57a40ba0ae851facdb8691cd5885e0a48b33c354b8c1723b7a5c5d6f0e`.
The baseline and candidate probe source hashes are respectively
`74f2271ca386e08515e50f5b0d87e9b4d6e514e7284084179114d2571e794a44`
and `65d0629300694e7bab1bbc15b053223e58f77db2832bf75523ac3ddf5b7615b4`.

The isolated run was pinned to CPU 29, used 31 samples per case, and ran
`before, after, after, before`. JSON and stderr captures are outside the
checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/expanded-run-a6a55b09f4-abba-20261021`;
each leg reported the exact preflight `base_parts=7 after_parts=8
base_bytes=5155 after_bytes=5320`. The table gives the range of the two leg
medians and the ratio of their mean medians, in microseconds:

| Case | Before leg range | After leg range | After / before |
|---|---:|---:|---:|
| `structural/commit` | 10.4140–10.4600 | 8.1720–8.8180 | 0.813931 |
| `structural/commit_save` | 44.8822–44.9622 | 43.1562–43.8422 | 0.968323 |
| `structural/forward_replay` | 72.8784–73.3264 | 69.6944–69.9564 | 0.955172 |
| `structural/inverse_replay` | 75.2126–75.5484 | 76.4944–76.5104 | 1.014883 |
| `history/retained128_build` | 4594.4050–4612.6560 | 2907.1160–2932.1270 | 0.634214 |

The dependency-disjoint optimization therefore removes the earlier commit and
forward-replay regressions and sharply reduces retained patch-history build
cost. Inverse replay remains a small 1.49% regression in this pair and is
retained as an owner-review item; this record does not claim that route is
accepted. The capture directory contains only external evidence, so no large
benchmark artifacts are committed.

## 9. Closure traversal repair

The pair above is rejected as an acceptance result because inverse replay was
1.49% slower. A one-sample Callgrind comparison of the same inverse route,
with the preflight retained and `LITCHI_PROBE_ONLY=structural/inverse_replay`,
reported 28,070,984 instructions for the baseline and 28,945,476 for the
candidate. In the candidate, `PartPatch::projection_unchanged` accounted for
842,395 instructions and its `closure_is_disjoint` walk for 836,047. The
capture and `Ir` annotations are outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/inverse-callgrind-a6a55b09f4`;
`sha256` in that directory records the files. This isolates the extra work to
the newly added proof rather than a second workbook projection parse.

The proof now walks only the recipient's before graph. A patch with no root
delta can change relationship owners only through its part deltas, so if none
of those parts is reachable before apply, no unchanged owner can make one
reachable afterward. The recipient-edge regression test proves the converse:
an unrelated inbound edge to a changed part forces eager projection
revalidation. Root edits, reachable deltas, malformed targets, and dangling
targets still use the full reparse path. Focused validation now passes 13 `opc`
tests and the full XLSB library suite passes 741 tests. The corrected paired
release ABBA and direct same-binary retained-history heap comparison are
recorded in §10.

## 10. Final closure-repair measurement and heap receipt

The closure-traversal repair is measured from commit
`3f085b89a6c3c0472bdfcc5287722b53ea2a98ee`. The baseline probe source hash
is `74f2271ca386e08515e50f5b0d87e9b4d6e514e7284084179114d2571e794a44` and
the candidate source hash is
`65d0629300694e7bab1bbc15b053223e58f77db2832bf75523ac3ddf5b7615b4`. The
release binaries were built into separate absolute target directories:

| Binary | Size | SHA-256 |
|---|---:|---|
| `tgt-expanded-before-3f085b89a6/release/w4b-expanded-probe-before-v2` | 7,643,696 bytes | `c38df593f0083aa3c0379d67363796adeb4f97076976393b8f6dddd61e5bce01` |
| `tgt-expanded-after-3f085b89a6/release/w4b-expanded-probe` | 7,489,816 bytes | `5cb9b6794318da9eaf19a57878b8c6d00a4b26d27cf66af2f555f3e3150efadb` |

The isolated native run is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/expanded-run-3f085b89a6-abba-20261021`.
It used `taskset -c 29 <binary> 31`, in the order
`before, after, after, before`; each route used 31 samples. Every leg reported
the same exact preflight: `base_parts=7 after_parts=8 base_bytes=5155
after_bytes=5320`. The table gives the range of the two leg medians and the
ratio of their mean medians, in microseconds:

| Case | Before leg range | After leg range | After / before |
|---|---:|---:|---:|
| `structural/commit` | 10.6180–10.6420 | 6.3620–6.5320 | 0.606491 |
| `structural/commit_save` | 44.6122–46.0542 | 41.9622–41.9744 | 0.925774 |
| `structural/forward_replay` | 72.5044–74.2204 | 67.4342–68.3904 | 0.925710 |
| `structural/inverse_replay` | 75.6704–76.6986 | 75.1044–76.1486 | 0.992676 |
| `history/retained128_build` | 4603.1260–4624.2160 | 2653.6150–2668.0050 | 0.576723 |

The capture manifest hashes are recorded in the external `binaries.sha256`
(`edd3306c0dd9206ef0a5258bfba838531952b9f088297803bfc762575ce0929f`) and
`captures.sha256`
(`a0340df3ef59c863b622b721f8ca60578773d70990e0ffa5be2e197b590825b6`).
These native results close the inverse-replay regression identified in §9.

The final retained-history heap receipt uses direct heaptrack invocation so
the debuggee is the probe itself:
`LITCHI_PROBE_ONLY=history/retained128_build taskset -c 29 heaptrack --output
<capture> <binary> 1`. Its external directory is
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/history-heap-3f085b89a6-direct`;
the four legs use the same `before, after, after, before` order. The
`heaptrack_print` K/M units below are decimal. Baseline legs report
594,916/594,920 allocation calls, 1,986/1,990 temporary allocations,
4.40M peak heap and 15.13M/14.98M peak RSS. Candidate legs report
210,153/210,154 allocation calls, 7,787/7,787 temporary allocations,
748.15K peak heap and 9.94M/10.03M peak RSS. Each leg reports 544B leaked.
Heaptrack timings are excluded from the native timing ratios above.

The earlier directory
`/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb/perf/history-heap-3f085b89a6`
is explicitly discarded: its command put `taskset` under heaptrack, and its
reports contain only `taskset`/locale allocations with no `litchi_xlsb` stack.
It contributes no heap claim. The direct captures in §4 and §6 invoke the
probe as heaptrack's debuggee and remain the cited receipts.

[0351]: 0351-api-contract.md
