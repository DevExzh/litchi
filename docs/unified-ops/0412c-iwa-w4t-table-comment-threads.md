# 0412c — W4-T: Pages and Keynote table comment threads

- Status: merge-ready on branch `uo2/iwa-w4t` after
  [0412b](0412b-iwa-w4t-numbers-cell-appearance.md).
- Date: 2026-11-02.
- `performance_claim`: every new route uses fewer instructions than its
  host route except Keynote's reply-bearing root set (+2.36%, within the 5%
  bar; §3).

The Pages and Keynote `table.cell_comment` rows that W1-G left host-only
("root-only shared slice; reply-thread lifecycle remains deferred") are
migrated: 5 Pages and 5 Keynote rows.

## 1. The shared reply engine

`litchi_iwa_table::comments_native` gains the reply half of the native
comment engine. It is written over the shared structural working set, so
Pages and Keynote share one implementation.

- `rewrite_reply(source, selection, ReplyOperation, profile)` reproduces the
  host's copy-on-write reply writer (`numbers::editor::model`,
  `{add,set,remove}_cell_comment_reply_at_location`). The root storage is
  cloned exactly under a fresh identifier. On the clone, the reply
  references are appended to, retargeted or removed. The cell's
  comment-list entry moves to the clone under a new key. The source root
  is removed, together with the reply that a replacement or removal gives
  up. An appended reply is a new storage beside the root, written by the
  root writer with the planned author (`materialize_author`, which is now
  shared with the root create). A replaced reply is an exact clone with
  its text patched.
- Admission: the root must be uniquely owned, and every direct reply must
  be a single-payload leaf owned only by that root
  (`thread::direct_replies`). A nested, shared or multi-payload reply is a
  `Dependency` refusal, raised before anything is staged.
- `rewrite_root`'s clear now removes the whole thread, as the host's
  `remove_unreferenced_comment_graph` does. That means the root, its
  direct replies and any generated author left without a referrer. The
  deletion-grade admission (`validate_clear_storage`) checks each reply as
  strictly as the root. A root's header may list its replies (field 4);
  only their multiset is checked, because the host appends replies after
  the author.
- A reply-bearing root's text set was never a native limitation. The
  in-place text rewrite keeps the reply references. The facades' refusal
  of that case is deleted.

## 2. Facades

`litchi_pages::table::comments::Edit` and
`litchi_keynote::slide::table::comments::Edit` gain `add_reply(text)`,
`set_reply(index, text)` and `remove_reply(index)` (ordinals, never native
identifiers). They also gain the typed refusal `Error::ReplyNotFound {
index }`. `clear` now removes a reply-bearing thread. Rewriting a reply
with its own text is an exact no-op. Both facades keep the existing
exact-source patch, inverse and semantic readback. The expected thread
keeps the root's metadata and a replaced reply's timestamp and author; an
appended reply carries none.

## 3. Parity and performance

`crates/litchi-iwa/tests/parity/{pages,keynote}_comment_thread_cases.rs`
cover two sources: Pages uses the native reply-bearing
`body-table-comment-hidden-native-saved.pages`, and Keynote uses a
host-authored two-reply thread. The host's new reply carries the clock and
a random storage UUID, so every write case is tier B with those two
new-object fields masked. Both outputs reopen, the host and leaf thread
reads agree, and the leaf inverse restores the exact source.

| Test (`pages_body_table_comment_` / `keynote_slide_table_comment_`) | Tier |
|---|---|
| `set_with_replies_parity`, `clear_with_replies_parity` | B |
| `reply_add_parity`, `reply_set_parity`, `reply_remove_parity` | B |
| `reply_refusals_are_typed` (out-of-range ordinal, same-text no-op) | — |

Callgrind, release parity binary SHA-256
`22c8c92d5cb438bec9a466e25100c3bf34383d51c6e99dc291b27b589e7cad3d`, CPU 5,
`pages_cost_probes::{probe_host,probe_leaf}`, ABBA
(`*_comment_{reply_add,reply_set,reply_remove,thread_set,thread_clear}_cost_abba`):

| Route | Pages host Ir | Pages leaf Ir | Pages ratio | Keynote host Ir | Keynote leaf Ir | Keynote ratio |
|---|---:|---:|---:|---:|---:|---:|
| reply add | 1,517,801,059 | 115,054,775 | 0.076 | 80,203,472 | 41,246,292 | 0.514 |
| reply set | 1,927,302,954 | 115,049,674 | 0.060 | 93,321,212 | 41,165,066 | 0.441 |
| reply remove | 1,888,923,825 | 114,495,511 | 0.061 | 82,555,688 | 40,796,772 | 0.494 |
| reply-bearing root set | 1,308,532,462 | 77,812,802 | 0.059 | 32,878,653 | 33,654,519 | 1.024 |
| reply-bearing root clear | 1,454,357,792 | 156,257,508 | 0.107 | 76,693,547 | 46,046,731 | 0.600 |

Ledger: 5 Pages rows (`set_table_cell_comment`, `clear_table_cell_comment`,
`{add,set,remove}_table_cell_comment_reply`) and 5 Keynote rows (the
`slide_table` equivalents) migrate. Two native queue rows are added.

## 4. Numbers

The four Numbers rows (`set_cell_comment` on a reply-bearing root and
`{add,set,remove}_cell_comment_reply`) stay host-only in this slice. The
Numbers leaf keeps its own reply engine (`comments_reply_native`), which
admits only a single-member graph. It refuses a reply thread written by the
host's legacy writer on a native document, so the leaf is narrower than
the host on those routes. The follow-up moves the Numbers reply family
onto `rewrite_reply`. It also aligns the Numbers profile's preview and
save-token policy with the host's text set on a reply-bearing root, which
keeps the root previews and stamps only the storage component. Until that
lands, a Numbers clear routes through the shared engine and now accepts a
reply-bearing root; every `litchi-numbers` comment test passes.

## 5. Verification

The gate script `.coord/iwa-w4t/gates.sh` runs on the merge of trunk; its
result is reported in the hand-off.
