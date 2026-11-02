# 0412d — W4-T: Numbers comment threads on the shared reply engine

- Status: merge-ready on branch `uo2/iwa-w4t` after
  [0412c](0412c-iwa-w4t-table-comment-threads.md).
- Date: 2026-11-02.
- `performance_claim`: every Numbers route uses fewer instructions than its
  host route except the reply-bearing root set (+1.03%, within the 5% bar);
  the Pages and Keynote thread routes are unchanged (§3).

The four Numbers `table.cell_comment` rows that 0412c §4 left host-only are
migrated. Numbers now writes reply threads through the same engine as Pages
and Keynote.

## 1. Numbers on `rewrite_reply`

- `litchi_numbers` direct-reply commits (`Package::{add,set,remove}_table_cell_comment_reply`
  and `edit_table_cell_comment_replies`) call
  `litchi_iwa_table::comments_native::rewrite_reply` with
  `Profile::NUMBERS`. The private single-member engine
  `package/comments_reply_native.rs` (3,423 lines) is deleted, along with
  its metadata/ZIP publication in `comments_reply.rs`. The private engine
  refused a reply thread that the host's legacy writer had added to a
  native workbook, so the leaf was narrower than the host.
- Numbers keeps its own admission ahead of the engine:
  - strict package metadata (`comments_metadata::strict_source`, `inspect`,
    `reject_unknown_archive_metadata`);
  - a root-owned, singly stored list entry;
  - a registry UUID binding for the root and for every direct reply the
    write may release.
- An appended reply is written by the package's commenting author and
  carries a creation date, as the host's writer does. The old engine copied
  the first reply's timestamp and author.
  `reply_text_edits_preserve_metadata_and_append_uses_the_commenting_author`
  (renamed) pins the new rule. A replaced reply keeps its own metadata.
- Setting the text of a reply-bearing root now goes through the shared root
  rewrite, which keeps the replies. Applying that patch is also admitted for
  a text replacement. Only the apply of a clear still requires a reply-free
  root.

## 2. Engine changes found by the Numbers suites

- **Shared roots.** A root several cells share (list refcount above 1)
  survives for the other cells. The entry's refcount is decremented (the
  host's `decrement_table_data_list_entry`), the cell moves to the copy, and
  nothing is removed. The copy, and a replaced reply's copy, take a fresh
  storage UUID, because otherwise two live storages would share one.
- **Mirrored headers.** When a root's header mirrors each reply in a
  `[4, ordinal]` FieldInfo, the mirror is kept complete and in payload
  order: append adds one, replace retargets it, remove drops it and closes
  the ordinal gap. A header without mirrors keeps the host's shape.

All 31 `litchi-numbers` reply tests pass, as do every `litchi-numbers` test
and the host's comment unit tests.

## 3. Parity and performance

`crates/litchi-iwa/tests/parity/numbers_comment_thread_cases.rs` builds its
source from `comment-reader-native-created.numbers` B2, with one reply added
by the legacy host writer. The oracle is the legacy writer behind new
hidden `__parity_{add,set,remove}_table_cell_comment_reply_legacy` seams,
because `Draft`'s reply methods delegate to the leaf. The ledger records
these seams as retired, like the existing ones. The cases are tier B with
the new-storage date and UUID masked, root previews invalidated, and the
leaf's structural metadata, matching the existing Numbers root cases.

| Test (`numbers_table_cell_comment_`) | Tier |
|---|---|
| `set_with_replies_parity` | B |
| `reply_add_parity`, `reply_set_parity`, `reply_remove_parity` | B |

Callgrind, release parity binary SHA-256
`6e689f3d4e89f6c1686cb7eb5af20b20eb2a80ac6e9db9f4fdd1aa1c1165be05`. It
was built before the Numbers cases' tier/evidence arguments moved into the
test bodies; the routes are identical. CPU 5, ABBA:

| Route | Host Ir | Leaf Ir | Leaf/host |
|---|---:|---:|---:|
| Numbers reply add | 274,115,383 | 176,953,261 | 0.646 |
| Numbers reply set | 292,312,257 | 176,872,921 | 0.605 |
| Numbers reply remove | 258,902,883 | 176,352,381 | 0.681 |
| Numbers reply-bearing root set | 73,465,058 | 74,222,759 | 1.010 |

On the same binary, the ten 0412c Pages and Keynote probes re-measure
within 0.1 percentage points of their 0412c ratios. Keynote's
reply-bearing root set is the highest at 1.024.

Ledger: `set_cell_comment` and `{add,set,remove}_cell_comment_reply`
migrate. One native queue row.

## 4. Verification

The gate script `.coord/iwa-w4t/gates.sh` runs on the merge of trunk; its
result is reported in the hand-off.
