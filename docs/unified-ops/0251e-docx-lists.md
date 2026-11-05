# 0251e — DOCX list definitions and `listStyle`

Plan step 7a item 6 of [0248](0248-coordinator-restart-plan.md) §6 Phase B.
This ports `crates/litchi-docx/src/numbering/lists.rs` from the stale
commit `4dba739fac` (L-DOCX-W10 draft) onto the 0251b graph changes.

## Leaf

**`numbering::{ListKind, AddedList}`** and the splice functions
(`numbering/lists.rs`). They are ported from the draft with two changes:

- identifier allocation scans the part alone. The draft also consulted a
  typed collection, but the scan already sees every identifier at any
  depth;
- the splice keeps `CT_Numbering`'s order: definitions after the last
  `w:abstractNum`, instances after the last `w:num` and before
  `w:numIdMacAtCleanup`.

The template is OfficeCLI's `listStyle` definition. It has nine levels:

- bullets cycle through `•`, `◦` and `▪`;
- numbers cycle through `%N.` in decimal, lower letter and lower Roman;
- each level is indented 720 twips more, with a 360-twip hanging indent;
- Strict documents get `start`/`lvlJc=start`.

**Verbs on the eager package-bound `document::Edit`** (`document/owned/numbering.rs`):

- `add_list(kind, start) -> AddedList` stages the new definition and
  instance as a graph change (0251b):
  - an existing numbering part is replaced exactly;
  - otherwise `numbering.xml` and its relationship are created.

  A paragraph joins the list through `format_paragraph_at` with
  `properties::Numbering { id: num_id, level }`.
- `list_kind(num_id) -> Option<ListKind>` reads the staged numbering part:
  the list is `Bullet` when its level-0 format is `bullet`, and `Decimal`
  otherwise.

Singleton parts now take Word's plain names, when free: `comments.xml`
(0251c) and `numbering.xml`. `add_related_part` keeps numbering its names
from 1.

## Unified surface

`set <paragraph> listStyle=… [start=N]` (paragraph registry, set-only) works
on any paragraph: body, cell or content control, by path or by scoped query.

| Value | Effect |
|---|---|
| `bullet`, `unordered`, `ul` | joins the immediately preceding paragraph's list when it is a bullet list; otherwise adds a new bullet list |
| `ordered` and any other value (`decimal`, `ol`, …) | the same for numbered lists |
| `none`, `remove`, `clear` | removes the paragraph's numbering (`w:numPr`) |
| with `start=N` | always adds a new list, level 0 restarting at N |

The immediately preceding paragraph is the previous block of the same
container. A query that selects adjacent paragraphs joins them into one new
list, because each later target continues the list of the target before it.

The level comes from three places, in order:

1. `numLevel` in the same write;
2. the paragraph's existing level;
3. level 0.

The numbering is written in one formatting operation per paragraph.

**Owners and batching.**

- New definitions need the owning package, so a source-backed handle is
  materialized (0251b).
- `none` keeps the source owner.
- `listStyle` and `start` are barriers in coalesced formatting runs.

**Read-back.** It goes through `numId` and `numLevel`. `listStyle` itself is
not read: reading it needs the numbering part in the node view
(L-DOCX-R6). The former `listStyle` property gap is replaced by the
set-only entry.

**Capabilities.**

- paragraph `listStyle` and `start`: set, not get or add.
- Both DOCX capability goldens are re-blessed and the OfficeCLI delta is
  regenerated.

## OfficeCLI parity

| OfficeCLI paragraph | Litchi |
|---|---|
| set `listStyle` bullet/ordered/none, continuing the preceding same-kind list | same |
| set `start` | same (always a new list) |
| add `listStyle`, `start` | not integrated (set after add) |
| get `listStyle`, `numFmt` | missing (L-DOCX-R6) |
| `add /numbering` (`num`, `abstractNum`) | not integrated |

## Measurements

The harness and method are the same as 0251a. The case is
`list_style_bullet`:

- leaf: `add_list` plus one `format_paragraph_at`, then commit;
- unified: `set /body/p[151] listStyle=bullet`;
- eager owner; the preceding paragraph is not in a list.

| Trial | Leaf µs | Unified µs | Ratio |
|---|---:|---:|---:|
| 1 | 71.6 | 74.9 | 1.046 |
| 2 | 71.0 | 75.8 | 1.068 |
| 3 | 71.6 | 74.4 | 1.039 |

The first version wrote `numId` and `numLevel` as two operations, at 1.16.
Folding them into one formatting operation brought the ratio inside the
budget.

The machine was shared: load average was about 26 on 32 CPUs. One earlier
round of trials, run at a higher load, scattered from 0.98 to 1.42 and was
discarded. The kept trials are the three consecutive ones that followed.

Evidence: `results/0251e-docx-lists/trial-{1,2,3}.json`. Regenerate with:

```
DOCX_BENCH_CASE=list_style_bullet taskset -c 5 target/release/examples/docx_element_set_bench
```

## Verification

| Command | Exit |
|---|---|
| `cargo test -p litchi-docx --all-features` | 0 (`numbering::lists` splice-order unit test; `element_graph_edits` list creation, extension, readback, replay and exact inverse) |
| `cargo test -p litchi-ops --all-features` | 0 (`docx_unified_nested`: create, continue, restart, remove and query join, on both owners; `docx_unified_write`: setter liveness after reopen) |

## Not claimed

- `listStyle` read-back.
- `add` with `listStyle`.
- Numbering-part paths (`/numbering`).
- Level overrides (`w:lvlOverride`).
- Picture bullets.
- Continuing a list across an intervening paragraph without an explicit
  `numId`.
