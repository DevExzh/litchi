# 0326 — PPTX query regression, path-text parser, typed coercion, element read misses

- Status: implemented on `uo2/perf-r13`, cut from `feat/unified-ops-api` at
  `fc274041da`; trunk `16b800350b` (record 0325) merged at `4423bd2096`
  before the final measurement and gates.
- Date: 2026-10-17
- Input: records [0315](0315-query-walker-read-path-misses.md) §6 (the
  PPTX slowdown on trunk, the path-text parser of record
  [0314](0314-iwork-batch-equivalence.md), typed length coercion) and the
  read misses of records 0319–0324.
- Scope: `format/pptx` (`Item`, `TreeItem`, the eager slide read of
  `owner.rs` and `pptx.rs`), `litchi-core::selector` (`path.rs`,
  `store.rs`), `litchi-ops` `path/arg.rs` and `target.rs`, `props.rs` and
  `value/length.rs`, the note lookups of `format/xlsx/features.rs` and
  `format/xlsb/features.rs`, the style and level values of
  `format/docx/parts.rs`; the new example `path_parse_bench`, an `--only`
  option of `unified-pptx-elements`, and the bench-matrix parser.
- `performance_claim`: paired measurements in §2–§5. Release builds
  (`unified-query`, `unified-pptx-elements`, the DOCX/XLSX/XLSB examples,
  `path_parse_bench`) and `profiling` builds (the iWork and RTF examples),
  one pinned CPU (`taskset -c 30`) of the shared 32-core host at load
  3–5, processes in B A A B order, every harness itself an in-process
  ABBA. Instruction counts are callgrind with `--cache-sim=yes` (`perf` is
  blocked, `perf_event_paranoid=4`). Evidence: [results/0326/](results/0326/):
  `build.sh`, `measure.sh`, `binaries.sha256`, `bisect-callgrind.txt`,
  `callgrind-summary.txt` and the run set `abba/` (base = trunk
  `16b800350b` with this record's new example; after = the merged tree).

## 1. Summary

| item | before (base) | after |
|---|---|---|
| PPTX `pptx_large_bytes_all_shape` (µs; ratio to the borrowed leaf) | 2,404 / 1,910; 8.54 / 6.86 | 1,709 / 1,676; 6.13 / 6.06 |
| PPTX `pptx_large_bytes_all_shape_owned` (fair leaf) | 3.02 / 2.41 | 2.14 / 2.10 |
| PPTX node size (`size_of::<pptx::Node>`) | 176 B (168 before record 0321) | 144 B |
| path text `TargetArg::parse("/body/p[@paraId=…]")` | 2,841 Ir, 20 mispredicts; 120–122 ns | 1,304 Ir, 4; 61–67 ns |
| path text `TargetArg::parse("/body/p[1234]")` | 1,189 Ir, 11 mispredicts; 58–63 ns | 1,133 Ir, 3; 56–60 ns |
| Pages `batch_10000` / `batch_10000_parsed` | 1.478–1.489 / 1.307–1.316 | 1.466–1.495 / 1.278–1.282 (no change) |
| RTF `paragraph_layout_alignment_typed` (text route) | 1.298–1.301 (1.223–1.228) | 1.219–1.224 (1.221–1.236) |
| RTF layout `last_path_typed` (text route) | 1.165–1.177 (1.111–1.117) | 1.118–1.124 (1.121–1.125) |
| PPTX `pptx_equation_get_v1` | 1.297–1.315 | 0.766–0.785 |
| XLSX `comment_get` | 1.102–1.120 | 1.015–1.028 |
| XLSB `comment_get` | 1.239–1.248 | 1.144–1.159 |
| DOCX `style_get_format` / `paragraph_list_format` | 1.222–1.235 / 1.180–1.193 | 1.231–1.245 / 1.170–1.178 (no change) |
| DOC and PPT read misses (records 0324, 0323) | — | not changed (§5.5, §5.6) |
| bench matrix open misses | 98 | 101 (+3 newly tracked read rows; §6) |

Results are byte-identical: the query-corpus differential, corpus
fidelity, the CRUD checklist goldens and the capability and registry
goldens pass unchanged; nothing was re-blessed (§8).

## 2. PPTX query regression

### 2.1 Bisect

Record 0315 found `pptx_*_all_shape` 15–35 % slower on trunk than at
`30eed270b5`. The `unified-query` harness was built from `git archive` of
each first-parent trunk commit and the full-query routes counted with
callgrind inside the timed loop (`bisect-callgrind.txt`):

| commit | | `pptx_small_bytes_all_shape` ×5 | `pptx_large_bytes_all_shape` ×1 |
|---|---|---:|---:|
| `30eed270b5` | record 0315's base | 606,011 | 21,976,728 |
| `dea0630741` | trunk before the 0321 merge (records 0316–0320) | 605,150 | 21,971,814 |
| `ba7669242a` | the 0321 merge | **657,726** | **23,365,993** |
| `fc274041da` | trunk (0321–0324, 0315's walker step) | 606,630 | 21,587,983 |

Records 0316–0320 cost nothing; the 0321 merge added 6.3–8.7 %. On
`fc274041da` record 0315's walker step had hidden it in the instruction
count (the other PPTX routes fell 8 % there), not in time: one process each,
`pptx_large_bytes_all_shape` read 1,638–1,801 µs at `dea0630741` and
2,135–2,137 µs at `fc274041da`.

### 2.2 Cause

Per match, the walker moves the yielded node about ten times (adapter
iterator, walk step, match slot, the unified wrapper, the caller). Record
0321's `Item::Hyperlink(LinkItem)` holds an owner (`TreeItem` or
`TextItem`) plus an `Arc<Link>`, eight bytes more than a `TextItem`, so
`Item` grew from 144 to 152 bytes and the PPTX `Node` from 168 to 176.
Callgrind shows the two effects on the large route: `memcpy` 5.78 M →
6.44 M instructions, and a new mispredicted branch per match (20,503 for
20,500 shapes) at `match result` of the walker's step, on the
niche-encoded `Result<Option<Node>>` whose layout the larger enum changed.
The new `elements::walk_children` costs 15 instructions per parent and
nothing per child.

### 2.3 Fix

- The three record 0321 items that embed a tree or text item
  (`Column`, `Hyperlink`, `Equation`) are boxed; they are built only by
  queries and paths that name them.
- A `TreeItem` kept two lazy cells for a frame's chart part and SmartArt
  data model, 24 bytes each; a frame is one kind or the other, so one
  `frame` cell holds either. `TreeItem` 112 → 88 bytes, `TextItem`
  136 → 112, `Item` 152 → 120, `Node` 176 → 144.
- `format::pptx::tests::query_items_stay_as_small_as_a_text_item`
  asserts `Item` ≤ `TextItem` + one word, `TreeItem` ≤ 88 and `Node` ≤ 144.

### 2.4 Every PPTX query route (`abba/query-*.jsonl`)

Unified µs, base-1 / base-2 → after-1 / after-2, then the ratios.

| route | base µs | after µs | ratio base → after |
|---|---|---|---|
| `pptx_large_bytes_all_shape` | 2,404 / 1,910 | 1,709 / 1,676 | 8.54/6.86 → 6.13/6.06 |
| `pptx_large_bytes_all_shape_owned` | 2,389 / 1,901 | 1,712 / 1,676 | 3.02/2.41 → 2.14/2.10 |
| `pptx_large_bytes_all_shape_fenced` | 2,390 / 1,899 | 1,702 / 1,679 | 3.18/2.54 → 2.27/2.24 |
| `pptx_large_bytes_all_shape_width` | 2,597 / 2,901 | 2,364 / 2,295 | 9.35/10.56 → 8.56/8.37 |
| `pptx_large_bytes_all_shape_contains` | 1,520 / 1,629 | 1,261 / 1,271 | 5.09/5.53 → 4.24/4.30 |
| `pptx_large_bytes_limit10_shape_width` | 46.4 / 46.4 | 41.4 / 42.2 | 9.60/9.59 → 8.53/8.71 |
| `pptx_large_file_all_shape` | 4,996 / 4,680 | 4,576 / 4,311 | 11.98/11.08 → 11.03/10.27 |
| `pptx_large_file_all_shape_owned` / `_fenced` | 5,026 / 4,728 | 4,548 / 4,324 | 5.40/4.96 → 4.82/4.54; 1.43/1.34 → 1.32/1.21 |
| `pptx_large_file_all_shape_width` / `_contains` | 3,033 / 2,798; 1,969 / 1,742 | 2,576 / 2,550; 1,475 / 1,462 | 7.30/6.73 → 6.22/6.16; 4.53/3.98 → 3.36/3.31 |
| `pptx_small_bytes_all_shape` | 11.1 / 12.5 | 9.4 / 10.6 | 6.56/7.48 → 5.62/6.34 |
| `pptx_small_bytes_all_shape_width` / `_contains` | 13.8 / 14.5; 9.7 / 10.5 | 12.5 / 12.6; 7.8 / 8.1 | 8.32/8.75 → 7.55/7.57; 5.64/6.13 → 4.56/4.67 |
| `pptx_small_file_all_shape` (and `_owned`, `_fenced`) | 24.8 / 24.7 | 24.6 / 24.8 | 5.74/5.75 → 5.78/5.66 |
| `get_shape`, `first_shape_contains` (all four families) | — | — | within ±3 % or lower (e.g. bytes `first_shape_contains` 6.41/6.25 → 5.79/5.70) |

The large in-memory query is back at record 0315's base (1,636–1,806 µs
there) and the fair-leaf ratio below it (record 0315: 2.08–2.39 at the
base). Callgrind (`callgrind-summary.txt`): `pptx_large_bytes_all_shape`
22.67 M → 21.62 M (−4.6 %), `_contains` 22.78 M → 20.45 M (−10.2 %,
mispredicts 32,460 → 11,453), `pptx_small_bytes_all_shape` ×5 637.5 K →
606.3 K (−4.9 %). The base includes record 0325, which had raised
the large route from 21.59 M to 22.67 M.

The fair-leaf ratios stay above the 1.10 trigger, as they did before
record 0321 (record 0280's G1–G2 and record 0315 §6: the owned node and
the per-match source check, kept by owner decision).

## 3. Path-text parser

### 3.1 Change

`Path::parse` already parsed `/name[N]/…` in one pass (record 0302);
anything else, attribute predicates included, took the general parser.

- The fast path now also takes `[last()]`, `[key]`, `[@key]`,
  `[key=value]` and `[@key=value]` written without whitespace, quotes or
  escapes (keys of ASCII letters, digits, `.` and `_`; values of any byte
  but brackets, parentheses, whitespace, quotes, `*` and the operator
  bytes), several predicates per step, and builds `[N]` in place. Bytes
  are classified by one table lookup. Anything it does not take — valid
  or not — goes to the general parser, which therefore produces every
  error, offset and message as before.
- `Path::parse` is `#[inline]` with the general parser out of line;
  `Inline::push` and `OneOrMany::push` skip the drop glue of the
  placeholder they overwrite (which owns nothing).
- `has_query_syntax`'s pre-check (record 0302) is a branch-free table
  fold.

### 3.2 Differential test

`litchi_core::selector::path::plain_tests::path_parser_differential`
collects every quoted string that starts with `/` from the workspace's
`crates/` tree (sources, unit and integration tests, the OfficeCLI corpus,
the query-corpus, capability, registry and node-view goldens) and
`docs/unified-ops/results/0289/` (CRUD checklist, fidelity): 6,573
strings on the merged tree. Each is parsed whole and cut at every
character boundary (211,682 inputs, 190,689 of them on the fast path); `Path::parse` and the
general parser give equal paths, equal `Display`, and equal errors (kind,
offset and message). `plain_fast_path_matches_the_general_parser` adds
`@` and `=` to its exhaustive five-byte alphabet and 45 hand-written
accepted and refused forms (`[last( )]`, `[k==v]`, `[k=r'v']`, `[1a]`,
non-ASCII keys, the predicate limit, …).

### 3.3 Measurements

Per parse (`path_parse_bench`, new; callgrind instructions and
mispredicted branches, then median ns base-1/base-2 → after-1/after-2):

| path | `TargetArg::parse` Ir (Bcm) base → after | ns base → after |
|---|---|---|
| `/body/p[1234]` | 1,189 (11) → 1,133 (3) | 57.6/62.8 → 60.4/55.9 |
| `/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]` | 2,186 (24) → 2,044 (10) | 95.5/102.0 → 98.8/95.4 |
| `/body/p[@paraId=1A2B3C4D]` | 2,841 (20) → 1,304 (4) | 120.1/121.6 → 66.8/61.2 |
| `/slide[1]/shape[@id=4]/paragraph[2]` | 3,536 (21) → 1,692 (8) | 144.5/145.7 → 79.4/74.9 |
| `/body/p[last()]` | 2,242 (16) → 1,476 (7) | 97.6/101.8 → 74.2/68.6 |

`Path::parse` alone: `/body/p[1234]` 759 → 743 Ir, `[@paraId=…]` 2,329 →
880, `[@id=4]` 2,934 → 1,205. In the Pages batch (callgrind of
`pages_bench --only batch_1000`, 17,002 parses, profiling builds of
`91c393fc42` and `7a7ef7c8fe`, before the trunk merge) `TargetArg::parse`
costs 948 → 944 instructions and 9.9 → 3.9 mispredicts per item.

Record 0314's legs (`abba/pages-*.jsonl`, `keynote-*`, `numbers-*`):

| harness | scenario | base-1 / base-2 | after-1 / after-2 |
|---|---|---|---|
| Pages | `batch_10000` | 1.478 / 1.489 | 1.466 / 1.495 |
| Pages | `batch_10000_parsed` | 1.316 / 1.307 | 1.278 / 1.282 |
| Pages | `batch_1000` / `_parsed` | 1.085 / 1.088; 1.049 / 1.054 | 1.082 / 1.081; 1.054 / 1.050 |
| Keynote | `batch_1000` / `_parsed` | 1.069 / 1.068; 1.052 / 1.048 | 1.062 / 1.062; 1.051 / 1.052 |
| Numbers | `batch_1000` / `_owned` / `_parsed` | 1.315 / 1.299; 1.272 / 1.260; 1.158 / 1.164 | 1.292 / 1.314; 1.257 / 1.248; 1.145 / 1.157 |

The `/body/p[N]` text route is unchanged in time: the batch's path
strings were already on record 0302's fast path, and the mispredicts the
change removes cost no measurable time there. What remains of the 0.2
between `batch_10000` and `batch_10000_parsed` (about 60 ns an item) is
the parsed `Path` itself: 616 bytes (six inline 96-byte steps, a 72-byte
predicate per step), initialised, moved three times (267 instructions of
`memcpy`) and dropped per item. Shrinking it without boxing — attribute
predicates must stay allocation-free (`path_allocations`) — needs a
smaller predicate representation in `litchi-core::selector`, a public
type change (§7).

## 4. Typed coercion

`PropValue::coerce` wrote a typed `Length` as text and parsed the text
again (record 0315 §3.4). `Length::coerced_in(context)` now gives the
value the text read would give whenever that read is certain to succeed
with that exact value: positions below 2^50 EMU (where the unit text
reads back exactly), extents and line widths within 0..=`i32::MAX`, font
sizes that print in points (a whole number of hundredths of a point that
is not one of centimetres, which `Auto` prints first) from 0.5 to
4,000 pt. Anything else — word spacing, which rounds to twips, and every
value the text read would refuse — still goes through the text, so
registry validation and every error are the text route's.
`props::tests::typed_lengths_coerce_as_their_text` compares the typed and
text routes, value and error (code, message, suggestion), over about
100,000 lengths in all five contexts.

| route (`abba/unified-rtf-*`) | base-1 / base-2 | after-1 / after-2 |
|---|---|---|
| `paragraph_layout_alignment_typed` (text) | 1.301 / 1.298 (1.223 / 1.228) | 1.219 / 1.224 (1.236 / 1.221) |
| `run_compound_format_typed` separate (text) | 1.095 / 1.091 (1.098 / 1.093) | 1.085 / 1.079 (1.082 / 1.083) |
| `run_compound_format_typed` dense (text) | 1.117 / 1.101 (1.094 / 1.095) | 1.087 / 1.079 (1.091 / 1.097) |
| layout `last_path_typed` (text) | 1.177 / 1.165 (1.117 / 1.111) | 1.118 / 1.124 (1.125 / 1.121) |

Typed lengths now cost what text costs. The RTF single sets stay above
1.10 where they were (records 0310, 0315): the remaining cost is the RTF
adapter's, not coercion.

Colours are still coerced through their text: their parse depends on the
colour context (scheme names, `auto`, spreadsheet tints), and no route
measures them.

## 5. Element read misses

Each case is a warm `get` and its `format` against the leaf read the node
answers from (records 0319–0324's harnesses); ratios per process,
base-1 / base-2 → after-1 / after-2.

### 5.1 PPTX equation get (record 0321)

The eager owner read the slide part as a slide (`SlidePart::from_part`,
about 450 K instructions) before looking its scene up in the handle's
scene cache, so a `get` through a slide paid the part read twice, once for
the cached scene and once for the equations. The part's payload, which
keys the cache, is now taken first (`EagerSlidePart`) and the part is read
as a slide only on a miss; a cached scene was read from that very payload,
so the skipped read could not have failed.
`pptx_equation_get_v1` 1.297 / 1.315 → **0.766 / 0.785**. It is below 1
because the native leg also reads the presentation's slide catalog; every
other `unified-pptx-elements` route is unchanged (writes 1.000–1.106,
reads 0.06–1.00, as record 0321).

### 5.2 XLSX note get (record 0320)

A note `get` copied every note of the sheet, sorted the copies row-major
and kept one. It now sorts (cell, position) pairs and copies the selected
note only; the stable order is the listing's. `comment_get` (note 20 of
40) 1.120 / 1.102 → **1.028 / 1.015**.

### 5.3 XLSB note get (record 0322)

A note `get` labelled (`Label::cell`) every note before choosing. It now
sorts the records stably and labels the selected one. `comment_get`
1.248 / 1.239 → 1.144 / 1.159: still a miss, by 0.8 µs. The remaining
cost is path resolution and the node (the leaf leg also decodes and sorts
the 40 records).

`note_lookup_order` (new) checks, on XLSX and XLSB with notes added out of
row-major order, both owners, that `comment[N]` for every N,
`comment[last()]` and `comment` give the listing's node, path, text and
format, and that `comment[10]` is `path_not_found`.

### 5.4 DOCX style and list paragraph reads (record 0319)

Style and numbering-level own keys are now matched by their compile-time
position (`From::Own(u8)`) instead of by name, and a part's `font` key
reads `font.latin`'s index from a constant. Neither moved the reads:
`style_get_format` 1.222 / 1.235 → 1.231 / 1.245, `paragraph_list_format`
1.180 / 1.193 → 1.170 / 1.178. Callgrind of `style_get_format`: the leaf's
`paragraph_properties` and `run_properties` (about 97 K instructions) are
in both legs; the unified leg adds about 3 K for the path and about 12 K
for the `format` iteration (72 entries, the run and paragraph sourced
keys of a style, about 150 instructions each, owned strings for the
present text keys). The `format` member iterates every key, so the API
does not let a caller build only the keys it wants; a borrowed property
set cached on the node (so present text keys need no copy) is the
follow-up (§7).

### 5.5 DOC reads (record 0324)

Not changed. Callgrind of `style_get_format` (`doc_parts_bench`): the leaf
reads one style's character and paragraph properties (3.65 K
instructions); the unified leg adds the path and node (2.6 K) and the
47-key `format` iteration (5.4 K), with no repeated parse or copy to
remove. `section_get_format`'s leaf leg copies one struct (1 ns), so no
unified read can approach its ratio; the absolute costs (0.3–0.6 µs) are
the fixed per-call cost of a path, a node and a key table.

### 5.6 PPT reads (record 0323)

Not changed: `format/ppt*` belongs to the PPT agent (record 0323), and
its F12 follow-up (per-node caches of a slide's comments and a shape's
runs, and a run leg that builds the same pieces) is that owner's.

## 6. Bench matrix

`tools/unified_bench_matrix.py` gains the record 0326 parser (`p0326`: the
PPTX query routes, record 0314's iWork batch legs with its `_parsed` and
`_owned` routes, the RTF single-target routes, and the element reads it
re-measures) and FOLLOW_RANGE runs to 0326. The curated file gains the
read rows of records 0319 (two), 0320, 0321 and 0322 that this record
re-measures, so they are superseded here, and four annotations.
`--check` exits 0.

| | open misses |
|---|---:|
| matrix before this record | **98** |
| record 0315's PPTX full-query rows (16) and RTF single-target rows (6), superseded by this record's re-measurement, still above 1.10 | ±0 |
| read rows newly tracked (curated from records 0319–0322): XLSX note get and PPTX equation get, fixed here (meets) | ±0 |
| read rows newly tracked and still above 1.10: DOCX `style_get_format`, `paragraph_list_format`, XLSB `comment_get` | +3 |
| **matrix after this record** | **101** |

No row the matrix tracked before reaches its budget here: the PPTX
regression fix returns those rows to their pre-0321 ratios, which were
already above the 1.10 trigger (§2.4), and the RTF typed routes now equal
the text routes, which stay above it (§4). The read misses of records
0323 and 0324 are not in the matrix (§5.5, §5.6).

## 7. Gaps and follow-ups

- **Path size.** A parsed `Path` is 616 bytes; its initialisation, three
  moves and drop are the text route's remaining per-item cost (Pages
  `batch_10000` 1.47–1.50 against 1.28 parsed). A compact predicate
  representation in `litchi-core::selector` (a public type change, so a
  glossary row first) or an in-place parse into the caller's `PathArg`
  (a new public entry point) would remove most of it.
- **DOCX reads.** Cache the parsed property set on the node, so `format`
  borrows present text keys instead of copying them, and skip absent
  sourced keys by a presence mask.
- **XLSB note get** (1.14–1.16): the per-call path and node cost.
- **DOC reads** (record 0324) and **PPT reads** (record 0323 F12): as §5.5
  and §5.6.
- **Colour coercion** still goes through text (§4).
- **PPTX fair-leaf ratios** (2.10–2.27 in memory, 1.21–1.32 file fenced):
  the owned node and the per-match source check (record 0315 §6).

## 8. Gates

Merged tree (`4423bd2096` + `895245e210` + this record), build
environment per the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`,
the memory-capped cargo shim, one cargo command at a time). Every exit
status was checked (`(deleted-coord-artifact)`, not committed).

| Command | Exit |
|---|---:|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-core -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features F] --all-targets -- -D warnings`, F = none, docx, pptx, xlsx, xlsb, rtf, pages | 0 each |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-core -p litchi-ops` | 0 |
| `cargo test -p litchi-core --all-features` (316, includes `path_parser_differential`) | 0 |
| `cargo test -p litchi-ops --all-features` in chunks: `--lib` (929), `--doc` (210), the 147 test targets in five groups (1,333) — includes the query-corpus differential, corpus fidelity, CRUD checklist goldens, `path_allocations`, `note_lookup_order` | 0 each |
| `cargo test -p litchi-ops --all-features --test capabilities --test registry` (not re-blessed) | 0 |
| the same with `--no-default-features --features` every litchi-ops feature but `encryption` (not re-blessed) | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (not re-blessed) | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,313 bans, 252 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (283 tests) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo check --locked --offline --manifest-path tools/perf-baseline/{unified-query,unified-pptx-elements}/Cargo.toml --all-targets` | 0 each |
| `python3 tools/unified_bench_matrix.py --check` | 0 |
| `python3 tools/test_no_3rdparty_references.py` | 0 |

No capability, registry or OfficeCLI-visible behaviour changed, so the
delta table and the OfficeCLI corpus were not regenerated. No public name
was added or changed (the new `Length::coerced_in` and `EagerSlidePart`
are crate-private); no glossary row changes. No dependency was added.

## 9. Not claimed

- That the path-text change makes any batch faster in wall-clock time
  (§3.3: it does not, measurably).
- That the DOCX, DOC, XLSB or PPT reads meet the 1.10 trigger.
- Any change to the RTF ratios beyond the typed routes.
