# 0251c — DOCX comments: read, add and remove

Plan step 7a item 4 of [0248](0248-coordinator-restart-plan.md) §6 Phase B.
Builds on the package-graph changes of [0251b](0251b-docx-hyperlink-writes.md).

## Leaf

**Verbs on the eager package-bound `document::Edit`** (`document/owned/comments.rs`):

- `add_comment_at(paragraph_path, range, &NewComment) -> id`.
  - `range` is an inclusive pair of direct-run positions, or `None` for
    the whole paragraph.
  - The anchor is `w:commentRangeStart`/`w:commentRangeEnd` plus a
    reference run (`w:commentReference`), staged as an element
    replacement.
  - The comment entry is staged as a graph change. Without a comments part,
    a new one (`comments.xml` beside the main part, or the first free numbered name; 0251e)
    and its relationship are added. Otherwise the existing part is replaced
    exactly.
  - The identifier is one above every identifier used by the part and by
    the main document's comment markers.
- `remove_comment(id)` strips:
  - the part entry;
  - the range markers and reference runs in the main document;
  - the matching entries of `commentsExtended`, `commentsIds` and
    `commentsExtensible`, when present.
- `NewComment::{new(author, text), with_initials, with_date}`.
- Strict documents use the Strict relationship type and namespace.

Both verbs travel in the main-document `Patch`, so replay and exact inverse
apply (the inverse of the first add removes the part it created).
Source-backed owners refuse the graph change, as in 0251b.

**Reads.** `source_backed::Package::comments()` reads the comments part
without materializing the package, matching the eager
`Document::comments()`.

## Unified surface

**Nodes.** `comments` and `comment` were gaps (`L-DOCX-R5`) and are now read:

| Path | Node |
|---|---|
| `/comments` | the comments part; present only when the main part relates one, so root views of uncommented documents are unchanged |
| `/comments/comment[N]` | the N-th comment in part order |
| `/comments/comment[@commentId=X]` | by native identifier (raw path API only, like `sdtId`) |

A comment's properties are `id` (read), and `author`, `initials`, `date` and
`text`, which are read and add-only. Its view text is the comment text, one
line per paragraph.

**add.** `add <paragraph> comment` comments the whole paragraph, and
`add <run> comment` comments one direct run. Both work under body, cell and
content-control paragraphs.

- `author` is required (`missing_property` otherwise).
- `text`, `initials` and `date` are optional.
- The returned path is `/comments/comment[N]`.

**remove.** `remove /comments/comment[N]` deletes the comment and its
anchors.

**Owners.** Both verbs go through the eager editor (0251b), so a
source-backed handle is materialized once.

**Capabilities.**

- `comments` and `comment` gain get and query.
- `comment` gains add and remove.
- Both DOCX capability goldens are re-blessed and the OfficeCLI delta is
  regenerated.

## OfficeCLI parity

| OfficeCLI `comment` | Litchi |
|---|---|
| add `author`, `initials`, `date`, `text` | same |
| get `id`, `author`, `initials`, `date`, `text` | same |
| add `runStart`, `range` | the parent path selects the anchor (a paragraph, or one direct run) |
| set `author`, `initials`, `date`, `text` | not integrated (add-only) |
| `anchoredTo`, `parentId`, `done` | missing (modern-comment replies and resolution) |

## Measurements

The harness, machine and method are the same as 0251a. The case is
`comment_add_paragraph`:

- leaf: `add_comment_at` and commit;
- unified: `add /body/p[151] comment`;
- eager owner; 200-paragraph fixture without a comments part.

| Trial | Leaf µs | Unified µs | Ratio |
|---|---:|---:|---:|
| 1 | 191.2 | 200.4 | 1.048 |
| 2 | 190.4 | 199.1 | 1.046 |
| 3 | 191.2 | 199.6 | 1.044 |

The ratio is inside the 10% single-op budget. The unified side also reads
the committed comments part back to return the new comment's path.

Evidence: `results/0251c-docx-comments/trial-{1,2,3}.json`. Regenerate with:

```
DOCX_BENCH_CASE=comment_add_paragraph taskset -c 5 target/release/examples/docx_element_set_bench
```

## Verification

| Command | Exit |
|---|---|
| `cargo test -p litchi-docx --all-features` | 0 (`element_graph_edits` comment add into new and existing parts, removal, exact inverse) |
| `cargo test -p litchi-ops --all-features` | 0 (`docx_unified_nested` comment add, read, identity lookup, removal on both owners, reopened; capability tests) |
| `cargo clippy -p litchi-docx -p litchi-ops --all-targets --all-features -- -D warnings` | 0 |

## Not claimed

- Comment `set`.
- Replies, resolution and modern-comment authoring (`commentsExtended`
  and related parts are cleaned up on removal but never written).
- Anchors that span several paragraphs, or that cover hyperlink runs.
- After a comment is added, the paragraph holds a comment-reference run. The
  existing `wrapped-or-unproven-selection` gap of 0251a still refuses
  direct-body run `set` there.
