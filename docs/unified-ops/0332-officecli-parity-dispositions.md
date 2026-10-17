# 0332 — OfficeCLI parity dispositions and the v1 corpus subset (audit WP1 + WP13)

- Status: implemented on `uo2/dispositions`, cut from `feat/unified-ops-api`
  at `d761892de3` (record 0330 merge); trunk merged at `243aab0e0b`
  (record 0327, XLS/ODS worksheet features). The counts below are after
  that merge, which closed many XLS/ODS gap rows (one stale tier key,
  `xlsx:hyperlink`, was removed).
- Date: 2026-10-17
- Input: the definition-of-done audit of 2026-10-17, §1 items 2 and 6, §3,
  §4, WP1 and WP13; the goal's "§14's three acceptance layers pass for the
  v1 scope, with every OfficeCLI delta recorded as wontfix or deferred";
  design §13 (tiers) and §14 (acceptance).
- Owned: `tools/officecli_capability_delta.py`, `tools/officecli_corpus.py`,
  `tools/officecli_dispositions.json`, their unit tests
  (`tools/test_officecli_{capability_delta,corpus}.py`), and the generated
  reports `docs/unified-ops/officecli-capability-delta.md` (plus its new
  per-format pages) and `docs/unified-ops/officecli-corpus-deltas.md`.
- This is tooling and data only. No Rust code, registry golden or corpus
  outcome changed.
- `performance_claim`: none.

## 1. The dispositions file

`tools/officecli_dispositions.json` has four parts:

- `policy`: `p0Formats` (`docx`, `xlsx`, `pptx`) and
  `propertyBreadthTier` (`P1`);
- `tiers`: the design §13 tier of each OfficeCLI element
  (`<vocabulary>:<element>`), of selected properties
  (`<vocabulary>:<element>.<property>`) and of corpus classes
  (`corpus:<class>`), each with its basis in the design;
- `dispositions`: 47 entries with 129 keys, each with an `id`, a
  disposition, a `ref` and a `reason`;
- `corpusScripts`: the tier of every example script.

Keys are `fnmatch` patterns. The most specific match wins. A pattern that
names the property beats one that names only the element, and that beats
one that names only the format. Equal specificity goes to the earlier
entry. So a few format-wide entries (`delta:doc:*`, `delta:docx:*`) carry
the bulk, and named entries override them where the decision differs.

The disposition values:

| Value | Meaning | Required |
|---|---|---|
| `planned` | to be closed | `ref`: a work package (`WPn`) or record id |
| `deferred` | resolves to `deferred-P1` or `deferred-P2` by the row's tier | a reason or a ref |
| `deferred-P1`, `deferred-P2` | as above, stated explicitly; must equal the row's tier | a reason or a ref |
| `wontfix` | a deliberate difference from OfficeCLI | a reason |
| `out-of-scope` | not part of the unified surface | a reason |

Tier rules:

1. Design §13 tiers are P0, "also in v1" (`v1`: typed chart CRUD), P1 and
   P2. Features §13 does not list get a tier with a stated basis (for
   example, equations and text boxes are P1 as §12 special objects;
   sparklines and slide zoom are P2).
2. **The v1 boundary binds DOCX, XLSX and PPTX**, as the goal's
   definition of done says ("docx/xlsx/pptx additionally cover the P0
   surface of §13"). In the other formats, a P0 or v1 feature is P1.
   `policy.p0Formats` changes this in one place.
3. A property of a P0 or v1 element is formatting breadth (P1), unless the
   tier list names it. The P0-named properties are: metadata core and
   application keys, paragraph/table style and list application, comment,
   hyperlink and notes text, picture source, size, position and
   alternative text, worksheet visibility, cell value, formula and type,
   and the text of runs, cells, shapes and slides. This follows the
   audit's own reading of P0 (verbs and named items, not every OfficeCLI
   formatting key). A property is never in an earlier tier than its
   element.

## 2. Where dispositions apply

**Capability delta.** A gap row is one of these:

- an element whose status is `gap`, `missing`, `pending` or
  `out of scope`;
- a property whose read status is `gap`, `element gap`, `missing`,
  `pending` or `raw-only`, or whose write status is `gap`, `partial` or
  `pending`.

Keys are `delta:<format>:<element>[.<property>]`.

- The properties of a gap element share its disposition, unless an entry
  names that element and that property.
- A `missing` row that a disposition classifies is renamed **`unlisted`**:
  it is still absent from `capabilities()`, but now carries a recorded
  decision. `missing` remains the untriaged status, and `--check` refuses
  it.

**Corpus.** Each class has a key:

- a named class (a delta `D*`/`U*`/`H1` or a gap `G*`): `corpus:<id>`;
  G1 is split per verb, `corpus:G1|<verb>`;
- an element or property group: `corpus:<format>|<verb>|<element>|<refusal>|<property>`.

A group whose element or property OfficeCLI's help schema names inherits
that capability-delta row's disposition (shown as `via delta`). It keeps
its own corpus entry only when the entry names the property literally. So
one decision covers both reports, and a corpus entry that only repeats the
delta becomes stale.

## 3. `--check`

Both tools still regenerate in memory and compare. They now also fail on:

- a `missing` (untriaged) delta row;
- a gap row or corpus class without a disposition, or without a tier;
- a P0 or v1 row that is deferred or wontfix;
- a deferred row whose tier is not its row's tier;
- a recorded corpus delta that is not wontfix or out-of-scope;
- a row whose follow-up names only a closed milestone (`M2 …`,
  `M1-extended (R4x)`, `R4x`, `M1-Δ1`) and whose disposition gives no
  current work-package or record id (WP13). A wontfix or out-of-scope
  reason supersedes the follow-up;
- any disposition, tier or script key that classifies nothing (stale);
- a corpus script without a tier.

Without `--check`, the generators still write the reports, print the
problems and exit 1. The coordinator's `rebless.sh` therefore keeps
working after a trunk merge, and the gate's `toolunit` step
(`test_the_committed_tables_pass_check`, `test_committed_table_is_current`)
turns red until the new rows are classified.

## 4. Result: the capability delta

Every gap row now has a disposition: 7,458 rows (elements and properties,
21 formats).

| Tier | planned | deferred-P1 | deferred-P2 | wontfix | out-of-scope |
|---|---:|---:|---:|---:|---:|
| P0 | 89 | 0 | 0 | 0 | 4 |
| v1 | 39 | 0 | 0 | 0 | 0 |
| P1 | 453 | 6,082 | 0 | 0 | 44 |
| P2 | 0 | 0 | 660 | 0 | 87 |

- The 54 untriaged element rows are now `unlisted` with a disposition.
- The 694 rows whose follow-up names only a milestone now carry a current
  id: a work package, the owning record (0322, 0323, 0324, 0327, 0328), or
  0332 for a deferral this record makes.
- **Planned, by work package:**
  - WP3 (PPTX P0 residue): paragraph/run/table-row insertion, comment and
    notes writes, picture set, slide text and layout writes.
  - WP4 (DOCX P0 residue): chart series and axes, table columns and the
    table/row/cell writes, comment set, picture set, style and list
    application.
  - WP5: the raw layer and `dump`.
  - WP10 and WP11: XLS and ODS element families and their properties.
    Property breadth of their supported elements is deferred (0327).
  - WP16: proposed in §7 below.
  - Record 0253 F3: DOCX application properties on the root.
- **Out-of-scope:**
  - OfficeCLI's `diagram`, in DOCX and PPTX: a Mermaid renderer (design
    §10), not SmartArt;
  - DOCX `fieldChar`/`instrText`: Litchi models a complex field as one
    `field` node;
  - native selection identities and numbering identifiers (ADR 0034,
    record 0319);
  - the `markdown` add-time format.
- **wontfix:** none in the delta. Every deliberate difference so far is a
  corpus delta (§5).
- The property tables move to
  `docs/unified-ops/officecli-capability-delta/<format>.md`, one page per
  format (11 pages, 75–115 KB each). The main table keeps the inputs, the
  statuses, the summaries, the dispositions legend and entries, and each
  format's element and contextual-gap tables (127 KB). With the new Tier
  and Disposition columns, the single file would have passed the 1 MiB
  limit (it was 868 KB).

## 5. Result: the corpus and the v1-scope subset

Every corpus class has a disposition:

- Deltas D1/D1b/D4/D4b/D5/D7/D8/D9/D10/D11 are wontfix, each with its
  record. U1/U2 (upstream script errors) and H1 (harness environment) are
  out-of-scope.
- G1 is split per verb: `validate`, `raw-set`, `add-part` and `dump` are
  planned WP5.
- G2 (`add --from`) is deferred-P1, and G4 (tracked removal) is
  deferred-P2.
- G5 (OfficeCLI-only path forms) and G6 (value vocabulary) are
  deferred-P1.
- The 532 element and property groups are covered: 416 inherit from the
  delta, 15 are `(more)` lists (see below), and the rest have corpus
  entries.

**v1-scope subset.** It is every counted line of the P0 and v1 scripts:

- P0: the cell, workbook-metadata, document, field, numbering, paragraph,
  run, picture and table scripts, plus the PPTX picture, basic shape,
  typography, table and text-box scripts;
- v1: every chart script.

The P1 and P2 scripts are outside it: conditional formatting, validation,
sheet settings, sections, content controls, shapes, text boxes, equations,
connectors, effects, showcase decks, Mermaid diagrams, pivots, slicers,
sparklines, animations, transitions, 3D, media, OLE and revisions.

**Status: red.** The subset has 3,443 lines:

| State | Lines |
|---|---:|
| passed | 1,877 |
| accounted | 941 (wontfix 40, deferred-P1 886, out-of-scope 15) |
| open | 625 |

The open lines, by the first class that blocks each:

| Blocked by | Lines |
|---|---:|
| planned WP3 | 432 (mostly `ppt/tables/*` cascades after the refused `add tr`/`table` and `set tc`) |
| unverified | 83 |
| planned WP5 | 57 (`validate` and `raw-set` lines inside v1 scripts) |
| planned WP16 | 38 |
| planned WP4 | 15 |

The 83 unverified lines are warned calls whose skipped list ends `+N`. The
outcome's 48-character detail bound hides those names, so no disposition
can account for them. Follow-up F1 below covers this.

## 6. Other changes

- **G1 text.** OfficeCLI's side now names the raw layer and `dump`.
  Litchi's side says `import`, `evaluate`, `recalculate`, `merge` and
  `refresh` have landed (records 0266, 0274, 0279) and are not in the
  class. The classifier's verb list now includes those five verbs. No
  verdict changed: the corpus totals are identical.
- **Tool fix: a false element match.** The delta tool matched OfficeCLI's
  PPTX `diagram` to Litchi's `slide`, through the path form
  `/slide[N]/…`, and reported it `supported`. OfficeCLI's `diagram` is a
  Mermaid renderer, not a slide, so it is now unmatched like DOCX
  `fieldChar`/`instrText`/`ptab` (`UNMATCHED`). The audit's "docx
  `diagram` missing" is the same element: it is not SmartArt (Litchi's
  `smartart`, which the help schemas do not describe).
- The out-of-scope reason for `sort` no longer names a milestone.
- New unit tests:
  - delta (15): key specificity, duplicate keys, entry validation,
    milestone detection, tiers (breadth, the element floor, the v1
    boundary), `unlisted`, inheritance and its override, format-wide
    property keys, untriaged rows, tier mismatches, milestone follow-ups,
    rows without a tier, stale keys and `--check` on an empty file,
    rendering with property pages, the committed file being fully used,
    and the Mermaid unmatch;
  - corpus (6): new verbs, cascade roots, delta inheritance and literal
    override, `(more)`, G1 per verb, the problems, the v1 subset with a
    red and a green case, and `--check` on an empty file.

## 7. Decisions for review, and follow-ups

- **WP16 (proposed): XLSX P0/v1 residue and root-level writes outside WP3
  and WP4.** It covers:
  - XLSX rich-text runs (`run` element gap);
  - `add` of a worksheet cell;
  - the cell `type`;
  - picture source replacement;
  - chart series from cell ranges on another sheet;
  - routing OfficeCLI's root-level metadata writes (`set / title=…`) to
    `/metadata` for DOCX, XLSX and PPTX, and the DOCX/XLSX root settings
    writes.

  It needs a lane, or reassignment to WP3/WP4.
- **Root-level metadata is `planned`, not out-of-scope.** Litchi writes
  metadata through `/metadata` (records 0253, 0254). OfficeCLI's spelling
  fails today, and the fix is a routing change, so the honest disposition
  is planned.
- **The breadth rule** (properties of P0 elements are P1 unless named)
  decides most of the 6,082 deferred-P1 rows. Changing a property to P0 is
  one tier key. `--check` then forces its rows to be planned.
- **F1, the corpus detail bound.** Raise or remove the 48-character bound
  on skipped-property lists in `crates/litchi-cli/tests/officecli_corpus.rs`
  (not owned here). That turns the 83 unverified lines into classifiable
  ones.
- **F2, registry naming.** The `unlisted` rows remain absent from
  `capabilities()`. The DoD wants every element gap listed there with a
  follow-up, so each format owner should add capability gap rows for the
  `unlisted` elements they keep (one line per element in `format/*/spec.rs`).
  The dispositions then become stale and are deleted.

## 8. Verification

See the final report of this branch for the gate commands and exit codes:
the Python unit tests, both `--check` runs (including against the
`3rdparty/OfficeCLI` tree the coordinator's rebless uses), the crate
boundary and non-iWork gates, `cargo fmt --check`, and the
`officecli_corpus` Rust test (not re-blessed).
