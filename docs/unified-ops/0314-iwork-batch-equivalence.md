# 0314 — iWork batches: equivalence with one-at-a-time application, batch costs, and the open element gaps

- Status: implemented on `uo2/iwork-batch`, cut from trunk
  `feat/unified-ops-api` at `75e5da7150`; trunk `30eed270b5` (records 0310,
  0313) merged at `26670389d2` (no conflict; it touched no iWork source).
- Date: 2026-10-16
- Inputs: the four batch scenarios that miss the PERFORMANCE BAR since
  record 0305 (Pages `batch_10000` 1.53×, Numbers `batch_1000` 1.32×, Pages
  `batch_1000` 1.08×, Keynote `batch_1000` 1.067×), the D-rule (a batch
  publishes what its items publish one at a time), then Numbers cell
  comments (request 19), Numbers row and column edits on tables with
  formulas, and Keynote and Pages table creation.
- Owner decisions applied: **D1**, **D2** (2026-10-15).
- Commits: `0bfe3ed0ba` (text-wire boundaries), `4d78cfaa19` (Pages
  paragraph commits), `cc8b214fc0` (Keynote boundary queries),
  `b3607fc251` (unified runs, property test, bench routes), `26670389d2`
  (trunk merge), and this record.
- `performance_claim`: paired leaf/unified medians in §3 (profiling builds
  of `--features pages,numbers,keynote`, one pinned core of a shared host,
  in-process ABBA, processes B A A B). Instruction counts are callgrind
  (`perf` is blocked). Evidence: [results/0314/](results/0314/)
  (`measure.sh`, `summary.txt`, raw outputs, `binaries.sha256`; binaries
  not committed).

## 1. Summary

| Item | Status |
|---|---|
| D-rule property test over the corpus (`tests/iwork_batch_equivalence.rs`) | **added**; it found violations in all three formats (§2) |
| Pages: several paragraph sets in one leaf commit dropped paragraph-style entries (leaf data loss) | **fixed** in `litchi-pages` (§2.2) |
| Keynote, Numbers: coalesced runs published other bytes or named another item than one-at-a-time application | **fixed** (§2.3, §2.4) |
| Pages `batch_10000` | 1.554 → 1.471; **miss** (§3, §4.1) |
| Numbers `batch_1000` | 1.310–1.335 → 1.301–1.318; **miss** (§4.2) |
| Pages `batch_1000` | 1.077–1.083 → 1.081–1.092; **miss** (§4.1) |
| Keynote `batch_1000` | 1.075–1.078 → 1.071–1.073; **miss** (§4.3) |
| Numbers cell comments (request 19) | **open**, diagnosed further (§5.1) |
| Numbers topology on tables with formulas | **open** (§5.2) |
| Keynote and Pages table creation | **not started** |

Every same-family run was already one leaf batch before this record (one
paragraph transaction per Pages section, one cell transaction per Numbers
table, one drawable commit plus one commit per text for Keynote): the
remaining batch cost is per item, in front of the leaf (§4). Reads and
existing single-operation paths are byte-identical; the query-corpus
differential, corpus fidelity, the CRUD checklist and the capability and
registry goldens pass unchanged (nothing re-blessed).

## 2. Batches publish what their items publish one at a time

### 2.1 The property test

`crates/litchi-ops/tests/iwork_batch_equivalence.rs` draws seeded
sequences of 1–9 sets per fixture over every Pages, Keynote and Numbers
package under `test-data/` (targets taken from each fixture's own
paragraphs, placeholders, shapes, pictures, notes, runs, slides, table
cells and stored and unstored Numbers cells; repeated targets; text, empty
text, multi-line text that refuses, frames, rotations, zero sizes that
refuse, visibility, bold, numbers, booleans, `null`, NaN, formulas). Each
sequence runs three ways — one `File::edit` closure, one atomic
`File::batch`, and a chain of `File::set` on successive handles — and the
three must publish the same bytes, or refuse with the same code at the same
item. Four sequences per fixture by default (about 1,200 sequences);
`LITCHI_BATCH_EQUIVALENCE_CASES` raises it. Before the fixes it failed in
all three formats (28 of 156 Numbers sequences, 10 of 220 Keynote, Pages on
the first multi-paragraph sequence); after them it passed with 60
sequences per fixture (2,340 Numbers, 3,300 Keynote, 2,340 Pages), and the
committed default runs in CI.

### 2.2 Pages: the leaf paragraph commit (`litchi-pages`)

`section::paragraphs::Edit::commit` published every staged change as one
splice from the section text before them to the text after them. A splice
that crosses paragraph breaks drops the attribute entries inside it, so
setting paragraphs 2 and 3 of `basic.pages` in one edit lost paragraph 3's
paragraph-style entry (field 5 at index 67), which committing the two sets
one at a time keeps. The leaf leg of `pages_bench` and the unified run
both took that path.

The commit now publishes exactly what committing each operation on its
own, in order, publishes:

- `litchi_iwa_text_wire::storage_boundaries_with_limits` reads where a
  storage's tables place their entries (one bounded scan of the root fields
  and table entries); `StorageBoundaries::admits_combined_splice(hull)` is
  `true` when the text is in one fragment, no index entry lies in
  `(hull.start, hull.end]`, no entry of a dropping or normalizing table
  lies at `hull.start`, and no overlapping range meets the hull. Its unit
  test checks that the admitted combined splice equals the sequence and
  that the refused one differs.
- Consecutive text sets share one splice while the paragraphs they touch
  (first to last, in the group's base text) are admitted, no set leaves a
  single-section storage empty before a later one (an emptied storage
  loses its text field, which a later splice appends last), and the group
  does not cancel out (a group ending at its base text replays). Insertion,
  removal, relocation and swap publish one transaction each.
- When every operation is a set, the whole sequence is checked at once
  without replaying it (the common case: one splice, as before). The 10,000-
  paragraph document of `pages_bench` holds entries at index 0 only, so its
  batches stay one transaction.

Test `several_operations_publish_what_they_publish_one_at_a_time` compares
both on every corpus section (all paragraphs, repeated, emptied, cancelling,
mixed with an insertion). `paragraph_operations_round_trip_on_a_plain_section`
now expects five transactions for its five structural operations on
`basic.pages` (it expected one).

The unified Pages run stages every set in item order (it used to sort and
keep the last set of a repeated line) and lets the leaf group them.

### 2.3 Keynote (`format/iwork/keynote/{write,coalesce}.rs`)

Before staging a set, a run now asks `Run::requires_barrier`; on `true` the
transaction publishes the run, plans the set again on the published handle
and starts a new run (`transaction.rs`, one hunk in the Keynote arm). A set
waits for the run when:

- it writes a text that already holds a set and the leaf does not admit one
  splice over the region the sets touch (`slide_text::Edit::admits_combined_splice`,
  `slide_notes::Edit::admits_combined_splice`,
  `slide_drawable::Edit::text_admits_combined_splice`, new; the admitted
  region is cached per text), or the staged text is empty and the base is
  not, or the set brings the text back to its base;
- it is a frame, slide-visibility or placeholder-visibility set that brings
  a staged value back to its base, or a second set of a staged table cell;
- it writes a text on a slide whose runs hold staged character properties,
  or character properties on a slide with a staged text (this was a typed
  refusal before; one at a time the two succeed).

The run also memoizes the paths it resolved (up to 16, compared by `Path`):
the benchmark and most batches set the same few elements again, and the
generic resolver cost about 4.5 K instructions per item.

### 2.4 Numbers (`format/iwork/numbers/coalesce.rs`)

- A formula publishes alone: its calculation objects take identities
  derived from the package it is written to (a value written first in the
  same transaction changed the object UUID in `Metadata.iwa`), and the
  leaf's formula budget is checked against that package
  (`formula-budget-native.numbers` accepted two formulas in one commit that
  it refuses one after the other).
- A new cell in a row before the rows of the cells the run creates starts
  a new run: the leaf creates new row headers in ascending order, one at a
  time in staging order (`chart-grid-replaced-native.numbers`, header
  bucket entries).
- A run whose commit fails finds the first failing item by committing ever
  shorter prefixes (bisection, on the error path only) and names it, as
  one-at-a-time application does; before, the error named the run's first
  item.

## 3. Measurements

Profiling builds of `--features pages,numbers,keynote` at `75e5da7150`
(base) and `b3607fc251` (after), `taskset -c 7`, B A A B (`measure.sh`):
Keynote and Pages `--rounds 8 --iterations 30`, Numbers
`LITCHI_NUMBERS_BENCH_SAMPLES=8`. Ratio is unified/leaf; the bar is 5 % or
20 µs, whichever is larger. Full table: `results/0314/summary.txt`; every
other scenario of the three harnesses is unchanged within noise.

| Format | Scenario | Base 1 / 4 | After 2 / 3 | Leaf / unified µs (after 2) | Bar |
|---|---|---:|---:|---:|---|
| Pages | `batch_10000` | 1.555 / 1.554 | 1.471 / 1.473 | 3207 / 4710 | miss |
| Pages | `batch_10000_parsed` (new) | — | 1.262 / 1.257 | 3210 / 4029 | miss |
| Pages | `batch_1000` | 1.083 / 1.077 | 1.081 / 1.092 | 2261 / 2452 | miss |
| Pages | `batch_1000_parsed` (new) | — | 1.050 / 1.052 | 2253 / 2364 | miss (by 0.2 %) |
| Keynote | `batch_1000` | 1.075 / 1.078 | 1.073 / 1.071 | 7870 / 8442 | miss |
| Keynote | `batch_1000_parsed` (new) | — | 1.046 / 1.047 | 7872 / 8236 | **met** |
| Numbers | `batch_1000` | 1.310 / 1.335 | 1.318 / 1.301 | 954 / 1257 | miss |
| Numbers | `batch_1000_owned` (new) | — | 1.254 / 1.231 | 951 / 1194 | miss |
| Numbers | `batch_1000_parsed` (new) | — | 1.123 / 1.164 | 951 / 1069 | miss |

New routes (new names only; existing routes unchanged): `*_parsed` pass
each path parsed once outside the clock (`TargetArg::parse`, then
`&TargetArg`, which the ergonomics bar offers for hot loops);
`batch_1000_owned` copies the `BatchItem`s outside the clock
(`File::batch` takes them by value; the existing route copies 1,000 items
inside it). Numbers outputs are compared byte for byte before timing
(`outputs_identical: true` for every pair); the Keynote and Pages
harnesses assert equality before timing.

The leaf legs changed only where §2.2 changed the leaf: Pages
`batch_10000` leaf 3091–3105 µs → 3207–3211 µs (+3.5 %: the hull scan and
the boundary read of the exactness check, on a 170 KB section text).

## 4. What the remaining per-item cost is

Callgrind of the base binaries (unified leg minus leaf leg, per item):

### 4.1 Pages

About 2.3 K instructions and 12 conditional-branch mispredictions per
`tx.set("/body/p[N]", ..)`: parsing the path text 0.96 K and 11
mispredictions (`TargetArg::parse` → `PathArg::parse` →
`litchi_core::selector::Path::parse`, `has_query_syntax`), five moves of
the parsed `Path` (inline step storage, 89 instructions of `memcpy` each),
the owned paragraph text, the plan and the staging. The leaf's own cost per
set is about 0.09 µs on the 10,000-paragraph document (its batch of 10,000
costs 0.8 ms over its batch of 1), so the 5 % arm leaves about 15 ns per
item at 10,000 items — less than the path parse alone. The `_parsed`
routes take the parse out: 1.05 at 1,000, 1.26 at 10,000. The rest is the
transaction frame per item; what remains of it is shared code
(`transaction.rs`, `target.rs`, `litchi-core::selector`), outside this
record's scope. The goal's own batch budget (a 1k batch within 2× of the
leaf, linear scaling) is met.

### 4.2 Numbers

4.7 K instructions and 22 mispredictions per item: the harness's copy of
1,000 `BatchItem`s inside the clock (about 1.2 K with the drop), batch
field admission 0.4 K, path parse 1.25 K, the cell plan (A1 parse, table
cache, stored-cell lookup, value coercion) and staging 1.5 K. Removing the
copy gives 1.23–1.25, removing the parse too 1.12–1.16. The leaf
publishes 1,000 cells in about 0.95 ms; the 5 % arm leaves 48 ns per item.

### 4.3 Keynote

Before: about 9.5 K instructions per item (generic path resolution 4.5 K,
path parse 2.65 K, staging 1.8 K, tree checks). The path memo (§2.3)
removed the resolution; the exactness check added about 1 K (a repeated
text compares its paragraphs with the base for the cancel rule, and a
repeated frame parses its length once more). With paths parsed outside
the clock the batch meets the bar (1.046–1.047); with path text it stays
at 1.07.

## 5. Not done, with what was found

### 5.1 Numbers cell comments (request 19)

A probe of `set_table_cell_comment` on every corpus workbook (scratch, not
committed) confirms record 0312: 31 files stop at `creation_graph`'s
package-wide census (more than one comment list), 6 at
`validate_comment_message_metadata`'s field-info rule, which mutation mode
applies to every list in the workbook. Native packages place each
`TST.TableDataList` in its own `Index/Tables/DataList-<id>.iwa`, the table
model in `Index/CalculationEngine.iwa`, tiles in `Index/Tables/Tile*.iwa`
and comment storages in `Index/Tables/DataList.iwa`. Creating a comment
on a table without a list therefore needs a new data-list member with its
registry entry and a reference from the model across members; on a table
with a list, a writer of two members (the storage's and the tile's). The
strict writer rewrites one member. Not changed; plan in
`IWORK_API_REQUESTS.md`.

### 5.2 Numbers row and column edits on tables with formulas

The reference implementation is the retired host editor's shift
(`litchi-iwa/src/numbers/editor/formula_dependency_shift{,/ast,/wire}.rs`,
about 4,000 lines over Prost archives: cell dependency records, range
precedents, formula ASTs, incoming references from other owners).
`litchi-numbers` rewrites exact wire payloads without Prost, so the port is
a rewrite of that logic onto its wire helpers. Not started;
`Unsupported::{Formulas, ReferencedByOtherOwners}` stay.

### 5.3 Keynote and Pages table creation

Not started.

### 5.4 Batch misses

All four batch scenarios still miss the 5 % arm (§3, §4).

## 6. Gates (merged tree `26670389d2` + this record, exit codes)

Build environment per the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`,
one cargo command at a time, the memory-capped cargo shim, test
executables pruned between chunks; driver `gates.sh`, not committed).

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa-text-wire -p litchi-pages -p litchi-keynote -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features keynote\|pages\|numbers] --all-targets -- -D warnings` (4 runs) | 0 |
| `cargo clippy -p litchi-iwa-text-wire -p litchi-pages -p litchi-keynote --all-targets -- -D warnings` (default features) | 0 |
| `cargo test --all-features -p`: litchi-iwa-text-wire 49, litchi-pages 697, litchi-keynote 1150 | 0 |
| `cargo test -p litchi-ops --all-features`, chunked: `--lib` 922, `--doc` 210, 139 test targets in 4 chunks (1256) | 0 |
| `cargo test -p litchi-ops --no-default-features --features` every feature but `encryption` `--test capabilities --test registry` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230); `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa-text-wire -p litchi-pages -p litchi-keynote -p litchi-ops -p litchi` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 228 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `LITCHI_BATCH_EQUIVALENCE_CASES=60 cargo test -p litchi-ops --all-features --test iwork_batch_equivalence` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
goldens, the delta table and the OfficeCLI corpus were not re-blessed or
regenerated. New public names (`StorageBoundaries`,
`storage_boundaries_with_limits`, `admits_combined_splice`,
`text_admits_combined_splice`) pass the glossary gate; no glossary row
changes.

## 7. Not claimed

- That any of the four batch scenarios meets the 5 % arm.
- That a combined splice equals the sequence outside what
  `admits_combined_splice` checks and the property test samples: the
  check is conservative by construction, and the property test is the
  evidence that the remaining rules (emptied texts, cancelling sets,
  formulas, new rows) cover the corpus.
- Native-app acceptance: this host has no macOS.
