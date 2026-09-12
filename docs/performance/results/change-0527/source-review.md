# 0527 XLSX row-primary-arena source review

status: bounded read-only review of the exact 0526 draft patch

production_change: none

This review checks the unapplied row-primary-arena draft against the fresh
baseline selected for 0527. It covers Rust ownership and type flow, every
constructor and consumer, scalar and formula payloads, unknown markup, limits,
and error ordering. It does not approve production retention or make a
performance claim. OLE2/OOXML remains the active priority; ODF and iWork are
outside this review.

## Identity and method

| item | identity |
| --- | --- |
| current baseline revision | `b6e0b05e6c01ce6205d4dc32a4c4cf37a2c7afda` (`b6e0b05e6`) |
| draft patch | `docs/performance/results/change-0526/row-primary-arena.patch` |
| draft patch SHA-256 | `28269e6051f9ee9453ccfd590a0f89269c400529d973a84f0ac908eb6730c5e9` |
| 0527 plan SHA-256 | `9762d905e3a197edd173d4c50ad508124a8a3577afcec25e5f179801c8781950` |
| 0527 baseline manifest SHA-256 | `0fab9bafc238611761659bcafdd2e7b5df2543aecb194ac5477c168a2d6e6096` |

The current baseline source hashes are:

| source file | SHA-256 | Git blob |
| --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` | `20a1b013e36eb8cd3ca21a5ec0bb44b9ef705ac7` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` | `3af9bd8c8a653bd690ac88effaa1e4888996accf` |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` | `c50324062526e4284cee4467f1dc38a8b3bf09f8` |

The three candidate source hashes recorded by the prior private-index review
are, respectively, `fea4369470a24600451bc96e7992e8a346b1c61a689ebc57d8e400da001ce736`,
`883ee4dcd9d848185401278665c631970712453bb6b5a73010228da3204a96a1`, and
`4d9841a90d7b76f946bc33f394abb9f35586179952873e9447c19c7d79a64064`.

The patch applies cleanly to the current baseline under `git apply --check`.
No patch hunk was applied. The review used source search, the exact diff, and
the pinned baseline; no Rust build, test, benchmark, profile, allocator run,
or capture was performed by this review.

## Static completeness result

The draft changes only `model.rs`, `scan.rs`, and `write/sheet_data.rs`.
There are no missed source consumers in the current tree:

* The two `RowSlot` constructors are covered: the empty-row constructor gets
  an empty `Box<[Span]>`, and the nonempty row close moves its pending row
  arena with `into_boxed_slice()`.
* The two `CellSlot` constructors are covered: an empty cell records
  `start..start`, and a nonempty cell records
  `PendingCell.primary_start..row.primary.len()`.
* The only pending-cell constructor records `primary_start`; the only
  pending-row constructor records a `Vec<Span>` arena.
* Both `write_cell` callers—and no others—pass `&row.primary`: the ordinary row
  writer and the value-only replacement-row writer. `copy_without` still
  receives a slice of the original `Span` type.
* `cell_slot`, expanded-dimension logic, package rewrite orchestration,
  merge rewriting, edit validation, and namespace-extension planning inspect
  cell or row metadata but do not consume primary spans. No other `.primary`
  consumer remains that would need a range lookup.

The new field types are internally consistent: `CellSlot.primary` is a
`Range<usize>`, `RowSlot.primary` is the owning `Box<[Span]>`, and
`primary.get(cell.primary.clone())` yields the `&[Span]` required by
`copy_without`. The split mutable borrows in the primary-close arm target the
disjoint `self.cell` and `self.row` fields. The patch has no unsafe code or
unchecked range access. A compiler run is still a required blocker before
the patch can be treated as mechanically verified.

## Scanner and payload preservation

The arena index is captured after the existing address and tag pipelines in
`start_cell`, so explicit and inferred coordinates, tag retention, attribute
decoding, and their error precedence are unchanged. An empty `<c>` uses the
current row arena length for both range endpoints.

Primary spans are appended at exactly the existing two storage events:

* a recognized SpreadsheetML `<f/>`, `<v/>`, or `<is/>` empty event appends its
  complete event span; and
* a recognized nonempty primary start records its `Frame.start`, then the
  matching `FrameKind::Primary` close appends the complete span.

The primary Start event itself does not append to the arena. This preserves
event order, repeated or mixed `f`/`v`/`is` children, and bytes between spans.
At cell close, all spans appended since `primary_start` belong to that cell;
at row close, the range indices remain valid because the arena is moved into
the same `RowSlot` that owns the cells.

Scalar payloads (`v` and `is`) keep their existing tag classification and
source spans. Formula payloads keep `scan_formula`, `formula_index`, text and
reference observation, and the shared/array/data-table dependency checks on
the existing event paths. The close arm still clears `formula_index` at the
same logical point. No formula metadata is inferred from, or removed with,
the arena representation.

Unknown direct cell children and MCE `AlternateContent` still set
`mce_payload`; they are not appended to the primary arena. The existing
validation refusal therefore remains in force. `extLst` remains opaque, and
when a payload is replaced, `copy_without` removes only the recorded primary
spans while copying intervening unknown bytes. Style-only updates, payload
clears with no new content, no-op paths, cell removal, and new-cell paths keep
their existing branches; only nonempty payload replacement resolves the range.

## Limits and error ordering

The event loop, `MAX_XML_EVENTS` check, `MAX_XML_DEPTH` check, quick-xml
end-name checking, source-position accounting, formula/resource observations,
and all scanner frame transitions are unchanged. The new arena has no new
unchecked arithmetic or index operation. Its length is bounded by the same
recognized primary events already admitted by the global event limit, although
its capacity and allocation pattern must be measured.

Reachable XML error ordering is preserved. Address parsing still precedes
`cell_tag`; formula attribute parsing still precedes the primary span append;
the cell close still checks cell state before row state; and row ordering is
checked before publication into `pending_sheet_data`. The patch adds a row
state check after the existing cell state check in the empty-primary branch and
in the nonempty-primary close arm. A cell can reach either arm only while its
row frame is active, so `cell.is_some()` with `row.is_none()` is an impossible
scanner state. The new error text is therefore unreachable through XML, but a
private-state test should still confirm a typed refusal if the test harness
constructs that state.

The checked writer lookup must remain in the existing nonempty,
payload-replacement branch. A malformed private range must return the typed
invalid-structure error and must not fall back to copying an unfiltered body.
The complete output validator, independent semantic readback, source identity
and cancellation fences, calculation-chain invalidation, and publication
checks remain outside this representation change.

## Concrete blockers before acceptance

1. Apply the patch in the isolated candidate checkout and run the compiler and
   focused tests. Static review found no missing constructor or consumer, but
   compilation is required to confirm the disjoint mutable borrows and all
   inferred `Range<usize>`/slice types.
2. Add differential writer coverage for empty and nonempty scalar cells,
   normal/shared/array/data-table formulas, duplicate and mixed primary
   children, both empty and start/end forms, and whitespace/comments/
   processing-instructions plus `extLst` between primary spans. Compare exact
   bytes and semantic readback with the baseline behavior.
3. Exercise unknown direct children and `AlternateContent` to prove the same
   typed markup-compatibility refusal and retained `mce_payload`; include
   style-only, clear, replacement, remove, new-cell, and no-op routes.
4. Retain the existing malformed-input checks for namespace aliases, inferred
   coordinates, duplicate/malformed cell attributes, mismatched end names,
   depth boundaries, and event-count boundaries. Add a private malformed-range
   writer test that proves no panic.
5. Complete the frozen 0527 native, allocator, and conditional profile gates
   on the unchanged baseline and this exact patch. Allocation-call reduction,
   fewer boxes, or a lower Callgrind edge is insufficient without useful
   end-to-end results and no rejected correctness/resource guard.

No source-level blocker requiring redesign was found. The patch is complete at
the constructor/consumer level, and its representation preserves the scalar,
formula, unknown-content, limit, and reachable error contracts by inspection.
Compiler/test confirmation and the frozen performance evidence remain concrete
acceptance blockers; production source must remain unchanged until they pass.

## Final applied-candidate review

The candidate is now applied in the working checkout for preflight review. The
candidate source manifest is
`docs/performance/results/change-0527/candidate/source-manifest.json`, SHA-256
`1d81835f8aa8827d38b21a0f730d1448f494d06049098d2eddda1c5381e13548`. Its
candidate source patch is
`docs/performance/results/change-0527/candidate/source.patch`, SHA-256
`870516943164482b7d09c499982251e159059223845666d32ff6f118a35500e4`.
The preflight-1 copies have the same two hashes. The changed-file manifest is
`source-diff.json`, SHA-256
`9cef678caff53ddc4313f1c21989d9cb7d44b78131cbdb7608b4e73402d10e0f`.

The three runtime production files exactly match the frozen draft hashes:

| production file | applied SHA-256 | frozen draft SHA-256 | match |
| --- | --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `fea4369470a24600451bc96e7992e8a346b1c61a689ebc57d8e400da001ce736` | `fea4369470a24600451bc96e7992e8a346b1c61a689ebc57d8e400da001ce736` | yes |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `883ee4dcd9d848185401278665c631970712453bb6b5a73010228da3204a96a1` | `883ee4dcd9d848185401278665c631970712453bb6b5a73010228da3204a96a1` | yes |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `4d9841a90d7b76f946bc33f394abb9f35586179952873e9447c19c7d79a64064` | `4d9841a90d7b76f946bc33f394abb9f35586179952873e9447c19c7d79a64064` | yes |

The remaining three changed crate files are test-only. The addition in
`cell_values/snapshot.rs` is inside `#[cfg(test)] mod row_reuse_tests`; the
other additions are in the `#[cfg(test)]` raw worksheet edit modules. Their
current hashes are `2f3f839adc91f0da204aefc83ebe2bb605cc02d269346759155b7bda54abc0e9`,
`c2b419115954b260c6c3aad4c0a04af4df66c7282c6b9aae1f9a48c3f1252166`, and
`0191f6a45fb9ee37d1fd71b0c7d87917787503c12a2ef4e7a01f174d1f3b7c3d`,
respectively. No release-path source file outside the three frozen
production files changed, so the applied candidate has no runtime extras.

The preflight receipt
`preflight-1/check-xlsx-tests.receipt.json` records the frozen all-feature
`litchi-xlsx` test command with exit code 0; its SHA-256 is
`764d2c8b1d3ae85b5dd86e2cc3336043607fd5302eec580a4b77329887c37a58`.
The output SHA-256 is
`326ace46a20427b39d90b73c528689185b2644871f53e4ee1b96cb4ad1610d8e`, and
the first test binary reports 986 passed, 0 failed. The six new focused tests
all report `ok`:

* the row-arena scan test checks exact per-row ranges for multiple `v`, `f`,
  and `is` spans, prefix aliases, empty cells, and empty rows;
* the writer test mutates a private range beyond its owning row and requires
  the typed invalid-structure error;
* the malformed-order test scans prior multi-primary content before each
  legacy malformed cell-attribute case;
* the ordinary/provenance test compares exact bytes while replacing duplicate
  and mixed primary spans with interleaved comment and opaque `extLst` bytes;
* the style-only test preserves duplicate primary content, opaque bytes, an
  empty cell, and an empty row; and
* the valid formula provenance test compares ordinary and provenance bytes and
  checks the complete semantic readback for a normal formula plus scalar cells.

These tests exercise the new range ownership and checked lookup without adding
production behavior. The release build and native/allocator evidence remain
separate acceptance gates; this final source review makes no performance claim.

## Final restore after pilot rejection

The row-primary-arena pilot is rejected for retention. The dense-r2 result
reported total p50 and mean reductions of `1.8214%` and `1.8105%`, respectively,
both below the frozen `2%` requirement. The commit and allocation gates passed,
but they do not satisfy the required end-to-end reduction gate. The arena
production change is therefore absent from the final checkout; no speedup claim
is made for it.

The final frozen source artifacts are:

| artifact | SHA-256 |
| --- | --- |
| `docs/performance/results/change-0527/final/source-manifest.json` | `9af673c4c13f2abb3aeaf5c4e31df6a613de6297abc104bf472931ff7733529f` |
| `docs/performance/results/change-0527/final/source.patch` | `1539d6673635c3220d05bc38c43f9e5254789499cbe5208c510aa66faa71bd88` |

The final manifest and patch restore the three production files to the exact
baseline hashes recorded above:

| production file | final SHA-256 | baseline match |
| --- | --- | --- |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs` | `71e23d40e8ef116ae833b51fb77c11067461fc6d5fb69f05fb2bfbe61644a0bd` | yes |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs` | `137a317c696e65043027007fbd04edf47f8407563bf0799b852c9be0d54996f8` | yes |
| `crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs` | `71bf1f126205c97b9ca964c1c6f927481becf2195833488470b2823266b426a5` | yes |

The final patch contains exactly the three test files named in the manifest.
The retained set has four tests: malformed-error ordering after prior primary
spans; ordinary/provenance writer byte equality for nonsemantic multi-primary
content; lossless style-only rewriting with duplicate primaries and empty
owners; and valid formula provenance omission with complete semantic
readback. They are all under `#[cfg(test)]` modules, so the restored checkout
has no runtime additions. The current test-only file hashes are
`2f3f839adc91f0da204aefc83ebe2bb605cc02d269346759155b7bda54abc0e9`,
`fd737a660a2cfcc7fc771f8b9db00af4e22e783123a368ead3deac8709d9fa91`, and
`0191f6a45fb9ee37d1fd71b0c7d87917787503c12a2ef4e7a01f174d1f3b7c3d`.

The two arena-only tests are absent from the final source. Their retained
artifact remains only as the un-applied
`docs/performance/results/change-0527/arena-only-tests.patch` (SHA-256
`8ebb971e50c7968356154e77cf0828c09f868f324fbc0ab6e7e619f5b64799d6`). The
four-test artifact used for the restore is
`docs/performance/results/change-0527/tests.patch` (SHA-256
`b9c3e20595217f5e5941d297744edbaa6c60a8c264f76fb906b19d13786d7c48`).

This bounded restore review performed no Rust job, build, test, benchmark,
capture, or other heavy tool run. The separate full12 final-quality run was
still in progress at review time; its result is not asserted here. The source
state is ready for the parent’s final-quality disposition, with OLE2/OOXML
remaining the active optimization priority and ODF deferred.
