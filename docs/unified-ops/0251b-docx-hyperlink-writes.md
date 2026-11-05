# 0251b — DOCX package-graph changes and hyperlink writes

Plan step 7a item 3 of [0248](0248-coordinator-restart-plan.md) §6 Phase B.
Follows [0251a](0251a-docx-element-writes.md).

## Leaf: graph changes in the main-document edit

A hyperlink to a URL needs a main-part relationship. This record adds the
same mechanism to the main-document transaction, which comments, pictures
and numbering need later.

**The operation.** `Operation::ChangeGraph { change, forward }` carries a
`document::GraphChange`. The change holds:

- relationships added to the main part;
- new parts related from the main part;
- exact replacements of existing part payloads.

The operation does not touch main-document XML. `Package::apply`, which
eager commits use, applies every staged change inside the same graph
candidate that publishes the main part. Preconditions are checked before any
mutation:

- forward: added relationships and parts must be absent, and each replaced
  part must hold its `before` bytes;
- inverse: the reverse.

A conflicting replay therefore fails without a partial publication.

**Patches.**

- `Patch::is_noop` is false when a patch carries a graph change, even if the
  main XML is unchanged.
- The inverse flips `forward`.
- The durable form is `graph.change`, with a blob in the existing
  transfer-graph encoding plus the replaced parts.
- `Package::apply_durable_document_patch` refuses such patches (it
  publishes main XML only).
- Source-backed owners refuse them, because their publication overlays the
  main document only.

**Verbs on the eager package-bound `document::Edit`** (`document/owned/graph.rs`):

- `add_external_relationship(type, target) -> rId`;
- `add_related_part(type, dir, stem, ext, content_type, payload) -> (name, rId)`,
  which picks the first unused name and id;
- `replace_part(name, payload)`;
- `part_payload(name)` and `related_part(type)`, which read the staged view;
- `insert_hyperlink_at(paragraph_path, at, HyperlinkTarget, &NewRun) -> position`;
- `set_hyperlink_target_at(hyperlink_path, HyperlinkTarget)`.

`HyperlinkTarget` is `Url(String)` (an external `hyperlink` relationship, of
the Strict type in Strict documents) or `Anchor(String)` (`w:anchor`).

**Markup.**

- A new hyperlink declares `xmlns:r` itself, so the reference resolves
  whatever the root binds, and carries `w:history="1"`.
- Retargeting rewrites only the start tag: it drops the old `r:id` or
  `w:anchor` and keeps every other attribute as written.
- The old relationship stays in the package. A later cleanup can remove it;
  that cleanup is not claimed here.

## Unified surface

**Paths.**

- `add /…/p[N]` with `hyperlink` or `link`;
- `set /…/p[N]/hyperlink[M]`.

Both work under body, cell and content-control paragraphs.

**add.**

- Exactly one of `url` and `anchor` is required (`missing_property`
  otherwise), plus `text`.
- The hyperlink's run takes the add-only formatting keys `font`, `size`,
  `bold`, `italic`, `color`, `underline` and `rStyle`.
- Without color, underline or style it gets OfficeCLI's fallback look: color
  `0563C1` and a single underline.
- The placement resolves against the paragraph's addressable runs, and must
  land beside a direct run.
- The returned path is `…/hyperlink[N]`.

**set.** `url` retargets the hyperlink, and `text` replaces its text
(0251a). Run keys on an existing hyperlink remain `unsupported_property`.
Format runs through `…/r[N]` instead.

**Owners.** Graph-changing verbs need the owning package. A source-backed
handle is materialized once (`to_owned_package_preserving_source`), and the
committed handle is eager.

**Capabilities.**

- `hyperlink` gains add (it already had set and remove).
- `url` becomes get, set and add.
- `anchor` and the seed formatting keys become get and add.
- Both DOCX capability goldens are re-blessed and the OfficeCLI delta is
  regenerated.

## OfficeCLI parity

| OfficeCLI `hyperlink` | Litchi |
|---|---|
| add `url`, `anchor`, `text`, `rStyle`, `color`, `font`, `size`, `underline`, `bold`, `italic` | same |
| add `tooltip` | not integrated (read only) |
| set `url`, `text` | same |
| default color theme `Hyperlink` or `0563C1`, single underline | `0563C1` fallback; the theme-linked style is used only when `rStyle` is given |

## Measurements

Same harness, machine and method as 0251a. The case is `hyperlink_add_url`:

- leaf `insert_hyperlink_at` and commit, against unified `add /body/p[151]`
  with a URL;
- eager owner; 200-paragraph fixture.

| Trial | Leaf µs | Unified µs | Ratio |
|---|---:|---:|---:|
| 1 | 157.3 | 163.2 | 1.037 |
| 2 | 156.2 | 162.7 | 1.041 |
| 3 | 156.7 | 161.8 | 1.032 |

Evidence: `results/0251b-docx-hyperlink-writes/trial-{1,2,3}.json`
(`DOCX_BENCH_CASE=hyperlink_add_url taskset -c 5
target/release/examples/docx_element_set_bench`).

The ratio is inside the 10% single-op budget. The leaf cost is dominated by
the main-document rescan of the element replacement.

## Verification

| Command | Exit |
|---|---|
| `cargo test -p litchi-docx --all-features` | 0 (`element_graph_edits`: replay, exact byte inverse, precondition conflict, anchors, retargeting, nested paragraphs, source refusal; `real_file_preservation` regains its changed-relationship case) |
| `cargo test -p litchi-ops --all-features` | 0 (`docx_unified_nested` hyperlink add and retarget on both owners, reopened) |
| `cargo clippy -p litchi-docx -p litchi-ops --all-targets --all-features -- -D warnings` | 0 |

## Not claimed

- No orphaned-relationship cleanup.
- No `tooltip`.
- Graph changes cannot be published on source-backed owners; those owners
  are materialized instead.
