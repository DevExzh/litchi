# 0759 — the spec-gap branch merged into the performance branch

Status: integration record, `performance_claim: none`. This record lists the
merge's inputs and every conflict. It also lists each judgement made where the
two sides disagreed, the code adapted so each side works with the other's
APIs, the gate results, and the failures that already exist on a tip. The
records it cites own their measurements; nothing here re-measures them.

OLE2 and OOXML stay the active priority, ODF optimization stays deferred, and
iWork crates were changed only where the merge itself required it (below).

## Inputs

| | Commit | Commits past the base |
|---|---|---:|
| Merge base | `f0ab67b55d` | — |
| Ours: `feat/office-format-completeness` | `e6cca92db2` | 449 |
| Theirs: `feat/spec-gap-implementation` (committed tip only) | `a67a38abf2` | 624 |
| Merge commit on `integration/spec-gap-merge` | `f592ecc1b0` (tree `c3c076acd0`) | — |

Only the committed tip `a67a38abf2` was merged. About 330 uncommitted changes
in the spec-gap worktree (`/home/zhuhe/code/litchi-spec-gaps`) belong to
another agent. They were neither read into the merge nor touched.
`feat/office-format-completeness` was not moved and nothing was pushed.

Excluding the incoming evidence directory, the merge changes 1,320 files
relative to ours. `docs/report/spec-gap-validation-evidence/` is kept whole:
203,490 files, 981.5 MiB.

## Method

- **Setup.** One trial merge (`merge.renameLimit=300000`) in a dedicated
  worktree, `/home/zhuhe/code/litchi-worktrees/merge-spec-gaps`.
- **Who resolved what.**
  - The coordinator resolved `soapberry-zip`, `litchi-core`, `litchi-opc`,
    `litchi-docx`, `litchi-ooxml-common`, the documents, the harness and the
    textual XLSX/XLSB/PPTX conflicts.
  - Four Opus subagents took the rest:
    - OLE2: `litchi-cfb`, `litchi-ole-common`, `litchi-doc`, and
      `litchi-ppt` verification;
    - ODG;
    - XLSX semantics;
    - XLSB and PPTX semantics.
  - Each subagent worked in the same worktree with its own build directory and
    checked its failures against the tips.
- **Pre-existing failures.** Every failure was classified as merge-caused or
  pre-existing. Pre-existing means it was reproduced on a detached checkout of
  that tip, with its own build directory: `tip-head` for `e6cca92db2`,
  `tip-inc` for `a67a38abf2`.
- **Union checks.**
  - `check_sides.py` reports lines either side added that are absent from a
    merged file.
  - The merged test inventory (`cargo test -- --list`) was compared name by
    name with both tips.
  - The pure auto-merge tree (`git merge-tree`, `b609e44091`) was diffed
    against the result, which isolates the 74 files adapted outside the
    conflicts ([list](results/change-0759/adapted-files.txt)).

## Conflicts and their resolution

There were 65 conflicts: 62 content, 2 add/add and 1 modify/delete. The
[full list](results/change-0759/conflicts.txt) is in the packet. Unless listed
under *Judgements*, a path keeps the union of both sides.

**Shared program logs.** These are `HOTSPOTS.md`, `REPORT.md`, `GOAL_AUDIT.md`,
`ADR_COMPLIANCE.md`, `BASELINE.md` and `CRUD_COVERAGE.md`.
- Ours added numbered sections, newest first. Theirs added unnumbered dated
  sections. Numbered sections stay first, then the incoming sections, at the
  insertion point both sides used.
- `GOAL_AUDIT.md` keeps our rewritten "Current evidence through 0502" table
  plus the incoming "Verified-cold per-file state evidence" row.
- No other text changed. `check_sides.py` finds no line either side added
  missing from any of the six logs.

**Record-number collisions.** The incoming side added
`changes/0470-performance-authority-reconciliation.md` and
`changes/0496-cold-verified-cache-post-observation.md`, next to our
`changes/0470-xlsx-empty-web-proof.md` and
`changes/0496-docx-edit-phase-attribution.md`. All four are kept; the slugs
tell them apart.

**`docs/adr/0024-current-topology.md`.**
- It keeps the union of both sides' amendments: our iWork amendments of
  2026-09-09 to 09-11, then the incoming XLDM amendment of 2026-09-11.
- A 2026-09-24 amendment states the merged inventory, computed from the merged
  `tools/crate_boundaries.json`: 65 packages, 244 declarations, 233 canonical
  edges, 12 development-only edges, debts 1, 2, 4, 8, 10 and 12–17, one host.
- `tools/test_check_crate_boundaries.py` pins 65 packages and 244 edges; all
  1,130 of its tests pass.

**`crates/litchi-iwa/examples/create_keynote_audio.rs`** (modify/delete). Our
deletion is kept. `76aee7e6ac` retired the host routes that example uses, and
the boundary checker forbids them there. The incoming change only moved the
example's imports, and nothing references it.

**Harness (`tools/perf-baseline`).**
- `README.md`: our 0508 section, then the incoming historical publication.
- `src/lib.rs`, hunks 1–9: our always-on operation metrics.
- `src/lib.rs`, hunk 10: 533 selectable cases (529 ours plus 4 incoming
  `Provider*`), 41 default cases, and the incoming assertion that `Provider*`
  stays out of the default.
- `src/bin/xlsb_crud.rs` and `src/corpus_manifest.rs`: both sides' lines.

**Crates.** Each conflicted source file is one of:
- a union of variants, exports, accessors and tests;
- adapted to the other side's API (see *Semantic adaptations*);
- decided as a judgement below.

The subagents' per-file notes are in
[resolution-notes.md](results/change-0759/resolution-notes.md).

## Judgements

The rule is from [0652](0652-owner-decisions-for-the-third-wave.md): where the
sides are incompatible, correctness and safety win, stricter validation and
typed refusals are kept, no `unsafe` is added, and accepted ADRs bind.

1. **soapberry-zip Deflate.**
   - Ours is kept: a fresh compressor per regenerated member.
   - The incoming `b769fef03a` carries one `DeflateWorkspace` across a plan.
     0618 measured and rejected exactly that (271 minor faults per publish), and
     ADR 0031's parallel deflate needs per-member independence.
   - Two incoming tests of the workspace itself (workspace selection, retry
     after a failed stream) are not carried.
   - Two incoming byte-identity tests (mixed payloads, empty flush) were
     adapted to our encoder.
2. **ZIP admission on every OPC ingress.**
   - The incoming admission applies everywhere, including ADR 0030's lazy path:
     an entry-count check before indexing, and refusal of encryption flags in
     central *and* local headers.
   - New `soapberry_zip::office::LocatedArchive` locates the archive with
     0632's single central-directory read, reports the declared entry count,
     then indexes. Both sides' properties hold.
   - Admission reads one 8-byte local-header prefix per entry, as the
     incoming side's security review specified.
   - Our 0623 read-count tests now remove exactly those probes, asserting one
     per entry, and keep their exact prefetch shapes (3, 9 and 13 reads).
3. **Header disagreement is refused at open.**
   - A central-directory CRC that disagrees with the local header or data
     descriptor is now refused at open. This is framing, with no
     decompression.
   - A payload that fails its CRC still surfaces at first decode, as ADR 0030
     requires.
   - Our fixtures that simulated payload corruption by flipping only the
     central CRC now flip it consistently in all three places, so they still
     test first-read refusal (XLSB threaded comments, PPTX master/layout).
   - Our OPC test of a disagreeing data descriptor now expects the refusal at
     open. Encrypted fixtures are refused at admission (OPC `source_xml_hint`,
     DOCX tail append).
4. **OPC provenance and publication.**
   - The incoming `PreservedRelationshipsXml {Source, Empty, Owned}` replaces
     our `CanonicalRelationshipsXml`. It publishes retained source bytes.
   - Our `Arc` pointer-pristine proof (0628/0647) stays on the merged type.
   - A part replaced after open loses pristine status.
   - Caller-supplied owned XML (incoming `try_replace_owned_xml_part`) is
     audited with our `VerifiedSource` before any mutation, because our
     provenance exempts retained bytes from the publication audit (0665).
   - The incoming source-backed readers now hold a counted `MonitoredReads`
     scope (0600) and disable read-ahead like every other publication read.
     Transfer monitors now live across prepare and publish.
5. **ADR 0030 decode-on-use.** Incoming callers that iterated
   `iter_parts()` for payloads now decode only what they read (`get_part`).
   - Whole-package size ceilings still charge inflated bytes through
     `try_iter_parts()`, on explicit Custom Data, pivot, SVG-lifecycle and
     InkAction operations only.
   - The incoming XLSX SVG census moved from `Edit::new` to the first SVG
     operation, so ordinary edits stay lazy.
   - Only `PartNotFound` means "absent" (0661). Incoming probes that read any
     error as absence were rewritten.
6. **CFB directory names.**
   - The incoming spec-correct case folding is kept; our ASCII fast path
     (0559) sits on top.
   - The length check now comes first (incoming order, bounded work). A name
     over 31 units that also holds `/` or NUL reports `TooLong` instead of the
     character error; it is still refused.
7. **CFB Reuse layout versus directory metadata.** Our Reuse copied the source
   directory image, which would silently drop edited timestamps and state bits.
   - If a caller supplies metadata that differs from the source, the writer
     rewrites in full, with new typed `SectorLayoutFallback::DirectoryMetadataChanged`.
   - Callers that supply none keep 0617's reuse; six real DOC files still
     reuse 6/6.
8. **DOC.**
   - The incoming FIB, protection and signed-source checks are kept. Our tests
     now use conforming fixtures (`with_valid_word97_dop`).
   - Our same-allocation `Patch::apply` shortcut also runs the incoming
     protection check.
   - A new `take_candidate` keeps 0730's retained render.
   - The 0753 fresh-writer goldens were re-pinned. The incoming Dop2002 fix
     (`581711a4c4`) pads every fresh DOP to 594 bytes, which changes those
     bytes, and the incoming tip emits the same nine digests.
9. **ODG.**
   - The incoming strict duration grammar is kept (XSD 1.0 2E §3.2.6.1); our
     tests now write `PT0.5S` and `PT1.0S`.
   - Our ODF 1.4 3D-scene owners (a scene inside `draw:g` or `dr3d:scene`) are
     kept, as are our child-order rules.
   - Our value batches run the incoming admission checks and fall back to the
     ordered path, so the error is identical.
10. **Namespace slices (0653) versus retained fragments.**
    - PPTX transition children re-declare inherited bindings only when the
      preprocessor rewrote the document. This matches the incoming `Raw`
      contract that two incoming tests pin.
    - Active transitions validate `mce::self_contained_fragment`.
    - DOCX revision paragraphs carry the incoming inherited
      `NamespaceBindings`. Our 0754 tracker scanner captures them
      (`NamespaceCapture::capture_tracker`, same bounds).
    - A splice that removes or inserts an `xmlns` declaration falls back to the
      full rescan.
11. **DOCX commit.**
    - Revision-disposition commits keep the incoming "publish projected bytes"
      semantics and skip our compaction policy.
    - They still run the UTF-8 validation that the compaction paths perform.
12. **XLSX verification order.** Our reduced readback (0744) runs first, then
    the incoming sparse verification, then the complete parse, whose result
    and error stay authoritative.
13. **XLSB.** The incoming drawing-load policy flows into our retained-parse
    publication: one validated parse, and the retained parse is published.
14. **MCE output growth.** Our doubling growth is kept. The incoming
    exact-reserve fallback is added for when the doubled target cannot be
    allocated. Growth is always capped at the output limit.
15. **Harness latency claim.**
    - The incoming rule gave up the latency claim whenever any allocation
      sample existed. Our normal binary records explicit *unavailable* samples,
      so its rows would have been labelled allocator-instrumented, which
      `perf_compare.py` rejects.
    - Only a sample the allocator wrapper actually observed now gives up the
      claim. A new unit test pins this.
16. **Full-default allocator contract.** The incoming 201-row contract covered
    its 37 default cases; our 0508 default has 41 cases and 213 rows.
    - The contract is now `default_case_matrix_213_rows`.
    - The four RTF/ODT/ODS/ODP `semantic_text_to_sink` runners open the same
      allocator region around `write_text_to`. Normal-binary output is
      unchanged.
    - The contract requires the plain RTF variant.
    - `FULL_DEFAULT_ALLOCATOR_COVERAGE.md` and the harness README were updated
      to match.
17. **CRUD coverage index v2.**
    - The incoming v2 binds the exact selector registry (443 names). It was
      re-bound to the merged 533, 90 of them ours; all 90 are excluded as
      `not-selected-in-representative-matrix`.
    - It takes 0508's catalog binding and conversion-export category, the only
      differences from the merged v1.
    - The v1 pin in `tools/test_crud_coverage_index.py` is now the merged v1.
    - [Generator](results/change-0759/scripts/regen_coverage_v2.py). v1 and v2
      both validate, and 50/50 tests pass.
18. **iWork example registration.** The incoming `97a97feb55` lists examples
    explicitly (`autoexamples = false`), which would silently stop building our
    `edit_slide_chart_data`, `edit_sheet_chart_data` and
    `edit_body_chart_data` (`fec3598f02`). They are registered under their
    existing names; nothing else in iWork changed.
19. **Fixtures broken by the other side's stricter checks.** The fixtures
    change; the assertions do not.
    - XLSX `drawing_anchor_geometry`: 0653 removed the per-element namespace
      expansion the test relied on. It now uses the expansion 0653 keeps (a
      declaration on a dropped `mc:AlternateContent`) and pins the same
      `LimitExceeded("output bytes")`.
    - XLSX `pivot_server_formats`: 0750's audit refuses to write `&#x1;`. The
      fixture restores those bytes with the raw archive writer, as our
      `a07d680852` did.
20. **Tracked `tools/perf-baseline/Cargo.lock`.**
    - It is refreshed offline for the merged manifests: `same-file` 1.0.6 and
      `winapi-util` 0.1.11 are added, at the root lock's versions. No existing
      version moved.
    - The incoming tip's harness and allocator gates fail `--locked` on the
      stale lock.

## Semantic adaptations

These make each side compile and behave against the other's changed APIs.

- **litchi-core.** The incoming `Reservation::try_merge` is adapted to our
  single-node reservation (`36bced7b26`).
- **litchi-ppt.** The incoming `picture_bullets.rs` uses our `RecordPayload`
  (0606).
- **litchi-cfb.** The incoming `create_stream_with_metadata` uses our
  `StreamPayload`, and a zero-copy `create_stream_shared_with_metadata` was
  added.
- **litchi-opc.**
  - The admission trait method `validate_unencrypted_entries` is implemented
    for our `SessionedArchive`.
  - Our deferred constructor retains content-types and relationship source XML.
  - The publication plan carries `PlannedXml::{Source, Authored}`.
- **litchi-docx.**
  - `XmlData::Shared` carries namespaces.
  - Ink, effects and SVG callers decode explicitly.
  - Transactions record `paragraph_namespaces`.
  - `ApplyRevision` string accounting was added.
- **litchi-xlsx, litchi-xlsb, litchi-pptx.**
  - Lazy-part callers, `SelectedPayload`, the placeholder-extension helpers now
    taking local names, and `ink_actions(&Presentation)`.
  - `workbook_projection` now includes `drawing_load_policy`.
- **Harness.** An incoming test passes our new `timed_docx_full_text` argument.

Every adapted file is listed in the packet.

## Gates

The 0675 runner ran with `CARGO_BUILD_JOBS=16` on the merge commit. Its logs
record `HEAD`, the index tree, and the unstaged and untracked counts (all 0).

| Gate | Result |
|---|---|
| fmt, check, clippy (`--lib`, `-D warnings`), rustdoc (`-D warnings`) | pass |
| tests (14 crates) | 13,221 passed, 0 failed, 78 ignored |
| facade | 382 passed, 7 ignored |
| facade-polyglot, allocator | 104 and 5 passed |
| claims (strict and structural), gate-tests, report, coverage, boundaries | pass |
| harness | 552 passed, **2 failed** (pre-existing), 1 ignored; the other targets pass 25 |
| non-iwork | **fails** (pre-existing) |

These were also run:

- `cargo test` for ODF, RTF, crypto, VBA and XLDM (13 crates): 5,794 passed,
  0 failed, 1 ignored.
- The facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt,ods,odp,rtf,encryption`:
  441 passed, 0 failed, 7 ignored.
- `cargo check --workspace --all-targets --keep-going`: the only failures are
  six pre-existing iWork examples. Non-iWork crates have no warnings.
- `cargo clippy --workspace --lib --no-deps -- -D warnings`: the only failure
  is the pre-existing facade `unit_arg` pair.
- `litchi-docx` clippy under every feature set passes: default, `encryption`,
  `fonts`, `automatic-fonts`, `vba-inspection`, `sign`, all features, and
  all-targets.

Logs are in [gates/](results/change-0759/gates/) and
[extra/](results/change-0759/extra/).

## Test inventory

Every test name on either tip is present in the merge, apart from:

- **14 gate crates.**
  - 38 `litchi-xlsx` `package::xldm` tests and 3 custom-data codec tests were
    moved by the incoming side into `litchi-xldm` and `litchi-ooxml-common`.
  - 13 base tests were replaced or renamed by the incoming side, which ours
    never changed.
  - 22 base tests were replaced or renamed by our records.
  - 2 incoming `DeflateWorkspace` tests were not carried (judgement 1).
- **ODF group.** Our `rejects_non_3d_dr3d_scene_children` is folded into the
  incoming `rejects_misplaced_dr3d_shape_owners`, which has the identical case.
- **Facade.** Three incoming names are tests that our `cb2a1a2d4f` renamed and
  0639 deleted with its dead function.
- **Harness.** The incoming side renamed one of our tests (see below).

Details: [inventory/](results/change-0759/inventory/README.md).

## Pre-existing failures

Each was reproduced on a tip; logs are in
[baseline/](results/change-0759/baseline/).

1. **non-iwork gate.** The failure is "workspace package inventory mismatch
   (unexpected: litchi-xldm)". The incoming side added `litchi-xldm` without
   registering it; identical on `tip-inc`.
2. **Harness `fresh_writer_corpora_are_deterministic_and_identify_the_packaged_stream`.**
   - The tiny DOC fresh-writer corpus is now `c9e22d55…`; the checked
     identity is `ec7824ca…`.
   - The cause is the incoming Dop2002 conformance fix (`581711a4c4`), which
     always writes the 594-byte DOP that nFibNew `0x0101` requires.
   - `tip-inc` fails identically once its stale lock is refreshed offline; its
     own gate could not build.
3. **Harness `pptx_native_image::tests::shapes_original_and_resaved_keep_their_typed_behavior`.**
   - This is the incoming rename of our `…_typed_refusals`.
   - The resaved fixture is refused, but with `picture blipFill stretch must
     contain one fillRect`, not the message the test expects.
   - It fails identically on `tip-inc`; ours passes on `tip-head`.
4. **iWork examples.** Six `litchi-iwa` chart examples,
   `create_{keynote,numbers,pages}_{donut,pie}_chart`, fail with E0308
   (`ChartData::new` now returns `DataError`). Identical on `tip-head`;
   `tip-inc` builds them.
5. **iWork test.** `litchi-numbers` test `table_relocation` does not compile
   (three dead functions under `-D unused`) on both tips.
6. **Facade lint.** Two `clippy::unit_arg` errors in the facade
   (`document/doc.rs:775`, `:952`) appear under workspace clippy on both tips.

None was fixed here:
- (1) is a gate-policy registration.
- (2) re-identifies a checked default corpus. It would re-pin the harness, the
  checked catalog, the V1 identity artifact and both coverage indexes, a
  measurement decision rather than a merge resolution.
- (3) needs the incoming owner to say which refusal message is intended.
- (4)–(6) belong to the tips' owners, and iWork is out of scope.

## Open risks and follow-ups

- **DOC fresh-writer identity (pre-existing 2).**
  - Until it is re-baselined, `doc_fresh_write_to` rows have a new corpus
    identity. They cannot be compared with pre-merge DOC fresh-writer
    measurements.
  - Owner decision: re-pin and regenerate the catalog, or keep the old bytes.
- **Admission cost on range sources.** Positional opens now make one small
  read per ZIP member. 0623 and 0632's other read savings stand. Coalescing
  the probes (for example, through the structural prefetch) is a candidate
  follow-up.
- **Decode cost where whole-package ceilings are counted.** Explicit Custom
  Data, pivot, SVG-lifecycle, InkAction, DOCX ink and effects inventories
  inflate every part. A no-decode declared size on `PartMetadata` would make
  them lazy.
- **Records measured before the merge.** Our records 0502 and 0504–0507 (ODG)
  and the OPC/DOCX open paths were measured before the incoming admission
  checks existed; their numbers describe the pre-merge code.
- **Incoming-side inconsistency.** `litchi-cfb` no longer folds
  supplementary-plane characters, but `litchi-ole-common`'s `cfb_path.rs`
  still does, and one test relies on it. Not changed here.
- **Stale sentence.** `CRUD_COVERAGE.md` still says the v2 index covers "all
  443 current selectors". That dated sentence is in a shared log and was left
  as written; the index now binds 533.

## Left out on purpose

- The spec-gap worktree's uncommitted changes.
- Pushing, and moving `feat/office-format-completeness`.
- Any change to iWork crates beyond the removed example and the three example
  registrations.
- Re-baselining the DOC fresh-writer corpus.
- Fixes to the pre-existing failures above.

## Cleanup

[cleanup.json](results/change-0759/cleanup.json) lists what was removed once
the gate logs were copied here: the two tip worktrees, every build directory
(about 614 GiB, plus about 112 GB the subagents removed themselves) and the
scratch directory. The integration worktree is kept.

Evidence and scripts: [results/change-0759/](results/change-0759/README.md).
