# Test inventory: both tips against the merge

The merge was checked name by name against both tips so that no test from
either side was dropped without a reason. Names come from
`cargo test --locked -p … -- --list` (script:
[`../scripts/list-tests.sh`](../scripts/list-tests.sh)). Each entry is keyed by
test binary and source file, with doctest line numbers removed.

- **Tip commits.** `tip-head` is `e6cca92db2` (ours) and `tip-inc` is
  `a67a38abf2` (theirs). Each is a detached checkout with its own build
  directory.
- **Gate-crate tree.** The merge side of the gate-crate listing ran on tree
  `c3c076acd0`, the tree of merge commit `f592ecc1b0`. The 14 gate crates are
  identical at every tree the gates ran on.
- **Other listings.** The facade names come from the
  `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt,ods,odp,rtf,encryption` runs. The
  harness names come from ours at 0756 (`../../change-0756/gates/harness.log`)
  and the merge's harness gate. The incoming harness could not build at its
  tip; see the record.

| Group | ours | theirs | merge | absent from merge |
|---|---:|---:|---:|---|
| 14 gate crates | 11,421 | 11,669 | 13,290 | 54 of ours, 24 of theirs |
| ODF, RTF, crypto, VBA (XLDM on theirs and merge) | 4,037 | 5,712 | 5,779 | 1 of ours, 0 of theirs |
| Facade (ODF features) | 448 | 331 | 448 | 0 of ours, 3 of theirs |
| Harness | 554 | — | 580 | 1 of ours |

## Why each absent test is absent

**Ours, 14 gate crates, 54**
([list](gate14-head-tests-absent-from-merge.txt)):
- **41 moved by the incoming side.** 38 `litchi-xlsx` `package::xldm` tests
  went to the new `litchi-xldm` crate, where one was split in two. 3 custom-data
  codec tests went to `litchi-ooxml-common/tests/custom_data.rs`. All run and
  pass there.
- **13 replaced or renamed by the incoming side.** Each existed at the merge
  base, and ours never changed it:
  - `711712282c`: the v4 range-lock sector tests;
  - `6bba8c3d48`: shared legacy RC4 in `litchi-crypto`, covering both
    `binary_rc4_secret_matches_apache_poi_vector` copies;
  - `f6e457de73`: the Word revision-metadata tests, 2;
  - `b58e42de1f`: `…exact_across_buffer_growth`;
  - `a521eaca28`: `…preserves_source_and_reopens`;
  - `9cc442c517`: the User Names stream, 2;
  - `141970fd6d`: Graph charts;
  - `ef88030948`: Morph transition edits;
  - `913eec261e`: the admission tests, 2.

**Theirs, 14 gate crates, 24**
([list](gate14-incoming-tests-absent-from-merge.txt)):
- **22 replaced or renamed by our records.** Each existed at the merge base,
  and the incoming side never changed it. The 13 commits:
  - `ef6dc6b68d`: overlay digests once per plan;
  - `69254a831f`: one source fence per shared read;
  - `36bced7b26`: budgets without per-level handles;
  - `37d4023794`: loosened eager publication audit;
  - `9556b46b4e`: DOC fresh-writer one-pass text;
  - `8662ca1bfd`: XLS SST order;
  - `d61133ae8b`: 0658;
  - `fadd5f68a2`: composed read-session budgets;
  - `cba0383f37`: BIFF globals in one pass (5 tests);
  - `554f323ef5`: windowed worksheet scan (2 tests);
  - `9445bbff5a`: 0667;
  - `44a4710699`: budgeted source-backed DOCX edits;
  - `5fa92d7ced`: 0657.
- **2 incoming `DeflateWorkspace` tests not carried** (record 0759, judgement
  1): `copy_and_store_preparation_leave_deflate_workspace_uninitialized` and
  `failed_reused_deflate_stream_is_discarded_before_retry`. Both exercise the
  workspace API, which the merge does not contain.

**Ours, ODF group, 1** ([list](odf-head-tests-absent-from-merge.txt)):
`rejects_non_3d_dr3d_scene_children`. Its exact case, a `draw:rect` child of a
`dr3d:scene` refused with `InvalidFormat`, is the second case of the merged
`rejects_misplaced_dr3d_shape_owners`.

**Theirs, facade, 3** ([list](facade-incoming-tests-absent-from-merge.txt)):
- Two ODT arbitration tests were renamed by our `cb2a1a2d4f`, with refreshed
  expectations (`…_does_not_hide_malformed_ooxml_catalog`).
- `refinement_restores_a_nonzero_cursor` was deleted by 0639 with the dead
  `refine_workbook_format` it tested.

**Ours, harness, 1** ([list](harness-head-tests-absent-from-merge.txt)):
`shapes_original_and_resaved_keep_their_typed_refusals` was renamed by the
incoming `fe57a1cd1a` to `…_typed_behavior`, which fails on both `tip-inc`
and the merge (record 0759, pre-existing failure 3).
