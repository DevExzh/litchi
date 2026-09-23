# 0744 — a strict `<sheetData>` lane and a reduced dense readback cut eager XLSX first-cell reads by 72% and one-cell commit+save by 60%

Status: retained, implemented in `crates/litchi-xlsx`. `performance_claim: none`
— the paired timings, instruction counts and allocation counts below are
evidence, not a registered claim. Output bytes are unchanged: every published
package, reopened cell dump and corpus archive is byte-identical between the two
legs (`results/change-0744/output-identity/`).

OLE2 and OOXML remain the active priority; ODF stays deferred and iWork is
excluded. Base `009d515bef`; branch `perf/0744-xlsx-eager-workbook-cell-path`,
commits `3166549994`, `5fb12908f0`, `6ac6b92d8a`, and after adversarial review
`3108e4dd40`, `b8367df8ca`, `a755ff8c26` (see [Review fixes](#review-fixes)).

## The redundant work

A coordinator sweep on the base showed `sheet.cell("A1")` on a fresh eager
`Workbook` materializing the whole 65,536-cell `dense-wide` sheet (29.1 ms,
about 440 ns per cell) and one-cell commit+save costing 152 ms (about 2.3 µs
per cell of the touched sheet). Frame-pointer profiles of the timed region
(matched `profiling` builds of both legs, `results/change-0744/profile/`)
attribute the base one-cell commit+save as follows:

| phase (dense-wide, one-cell commit+save) | owner | base share | complete pass over the touched sheet |
| --- | --- | ---: | --- |
| deflate of the changed sheet | OPC writer, `zlib-rs` | 24.46% | yes (necessary) |
| edit layout scan | `raw::worksheet::edit` `scan` | 19.16% | yes — namespace-resolving reader |
| base store parse | `Worksheet::store` → `raw::worksheet::parse` | 18.79% | yes — namespace-resolving reader |
| post-write verification parse | `raw::worksheet::parse` | 18.51% | yes — result discarded above 4,096 cells |
| compaction | `raw::compact::changed_worksheet` | 11.91% | yes — re-serializes every tag |
| publication audit | `xml_minifier::audit::verify_source` | 5.20% | yes (OPC writer) |

Four of the six complete passes were the same generic `quick_xml::NsReader`
work over the same benign bytes: per event, namespace push and resolution
(`NsReader::process_event` and `resolve_event` alone were 29% of a parse),
attribute iteration with duplicate checks, a decoded `String` per cell
reference, and dispatch through a dozen name comparisons. The untouched sheet
is never parsed, serialized or deflated — it is raw-copied by the preservation
writer (0.13% of samples) — so no cross-sheet redundancy exists; the
one-percent case costs twice the one-cell case because it touches both sheets.
No superlinear work was found in these paths: every pass is linear in the
touched sheet.

Namespace resolution is not needed inside a benign body: when the default
namespace in scope at `<sheetData>` is `SpreadsheetML` and no element in the
body declares a namespace, every unprefixed `row`, `c` and `v` resolves to that
one binding. That observation is what the lane proves once instead of per
event.

## What was changed

All production changes are private to `crates/litchi-xlsx`: no public API
change, no new dependency, limit, cache, thread or `unsafe`.

**1. A strict recognizer for the benign `<sheetData>` body**
(`src/raw/worksheet/lane.rs`, new). `recognize` and `walk` admit exactly

```text
body  := (ws | row)* "</" sheet-data-name ">"
row   := "<row" attrs "/>" | "<row" attrs ">" (ws | cell)* "</row>"
cell  := "<c" attrs "/>"   | "<c" attrs ">" ws? value? ws? "</c>"
value := "<v/>" | "<v>" text? "</v>"
attrs := (" " name "=\"" value-char* "\"")*   -- no xmlns, xmlns:*, xml:* or duplicate names
text  := any byte except "<" and "&"
```

where a name is ASCII (`name` or `prefix:name`) and a value contains none of
`"`, `<`, `>`, `&`, tab, CR or LF. For an admitted body `walk` reports, in
document order, exactly the `Start`, `Empty`, `End` and `Text` events
`quick_xml` reports for the same bytes, each tag carrying the bytes its
`BytesStart` would hold. Comments, processing instructions, CDATA, references,
formulas, inline strings, prefixed or namespace-declaring elements,
`xml:space`, irregular spacing or quoting, a second value, any other child and
truncated input all *decline*. A decline is not an error; the ordinary reader
keeps the whole body. Every pass runs `recognize` before replaying events, so a
decline never leaves partial state. `Entry::locate` takes the reader's
positions only after finding `<`, the start tag's exact name bytes and a
delimiter at them, and declines parts that begin with a UTF-8 byte-order mark,
whose reader positions exclude the mark.

**2. The worksheet parser takes the lane** (`src/raw/worksheet/codec.rs`).
`Parser::parse` drives the reader up to the root's `<sheetData>` start tag.
When its unprefixed children resolve to a `SpreadsheetML` namespace and the
body is admitted, rows go through the unchanged `start_row` and `finish_row`;
cells go through `cell_fields` — the former body of `start_cell`, now shared by
both routes, so coordinate, style and metadata checks and their order are one
piece of code — then `push_text`'s byte bound and the refusals `finish_cell`
can raise for the lane's closed cell shape, at the same events. Admitted cells
are kept as a compact `LaneCell` whose value text is borrowed from the document
(`PendingCell`/`RawCell`, 192 and 224 bytes, are skipped) and are materialized
through the unchanged `materialize` after the parse, in document order, so
materialization and shared-string refusals keep their precedence. The reader
then resumes over a spliced copy of the document without the body
(`lane::splice_without_body`, `lane::skip_to`): the body is balanced and
declares nothing, so the reader's nesting, namespace bindings and end-name
stack at `</sheetData>` are the ones the original reader would have had, and
`quick_xml` error texts carry no document offsets.

**3. The edit scanner takes the lane**
(`src/raw/worksheet/edit/codec/snapshot/scan.rs`). Rows go through the unchanged
`Scanner::start`, `empty` and `finish`; cells reproduce `start_cell`
(`cell_address` and `cell_tag`) for the admitted shape. Spliced reader
positions are shifted back to real offsets. The lane is admitted only when the
body is UTF-8 and its exact reader-event count fits the remaining event budget,
so the `MAX_XML_EVENTS` refusal is untouched. A scanned cell's single payload
span is now stored inline (`PrimarySpans`), removing one boxed slice per cell
on both scanner routes.

**4. Compaction takes the lane** (`src/raw/compact.rs`) at the element the parser
and scanner treat as the body: the first `SpreadsheetML` `sheetData` child of a
`SpreadsheetML` `worksheet` root whose unprefixed children resolve to
`SpreadsheetML`, decided once there. Every admitted tag is
already in the exact form `write_start` and the writer emit, so the compact body
is the source body with its formatting whitespace removed unless an ancestor
preserves it. The web-binding proof is told what per-event observation would
have concluded: inside a lane body only an apostrophe in an attribute value or
non-UTF-8 value text can end it (`Probe::decline`).

**5. Reduced readback for dense value edits**
(`src/workbook/edit/semantic/transaction.rs`). A cells-only eager commit whose
cell rewrite is the sheet's only byte change writes through the existing
`rewrite_value_only_with_provenance`, which emits exactly `rewrite`'s bytes
(existing tests and a direct randomized differential of the two writers: 3,000
generated worksheets and random cells-only plans, 1,123 through the provenance
writer, 920 refused identically). It verifies with change
[0525](changes/0525-xlsx-unchanged-cell-readback.md)'s `reduced_readback` when,
and only when: compaction left the rewrite unchanged, so the provenance spans
index the published bytes; compaction admitted the worksheet's own body — the
first `SpreadsheetML` `sheetData` child of a `SpreadsheetML` `worksheet` root
whose unprefixed children resolve to `SpreadsheetML` — to the lane; the part
does not mention the markup-compatibility namespace; the part is at most
16 MiB, the web-extension reader's limit, which every larger changed worksheet
fails at a later step anyway; and the complete store would exceed the
validated-store handoff bounds of change
[0025](changes/0025-xlsx-validated-store-handoff.md), so the reduced store never
reaches a snapshot. The reduced store is then refused, as 0525's merge refuses
it, when one of its cells falls inside an omitted range. Any refusal falls back
to the complete parse, whose result and error stay authoritative.

**6. Smaller per-cell work.** Number validation skips the `f64` parse for an
optionally signed run of at most 18 ASCII digits, which is always a finite
binary64 value (`src/cell.rs`, differential test against the float parser).

## Why it is sound

**Exactness of the lane.** The recognizer is a pure function of the bytes, and
for every admitted body the event sequence it reports equals the reader's:
tags are delimited by the same `>`, their content is the reader's
`BytesStart` content, text runs end at the same `<`, and no reference can split
them. Attribute values without `&`, tab, CR or LF are returned unchanged by
`quick_xml`'s decode-and-normalize, and duplicate names are declined rather
than reproduced. The default-namespace check at entry plus the ban on `xmlns*`
attributes make every body element resolve to the `SpreadsheetML` binding the
passes test for. Each pass applies the lane events through its own methods;
where a pass bypasses one (lane cells in the parser and scanner, tags in
compaction), the replacement reproduces that method's outcome for the only
shape the lane admits and is covered by differential tests.

**Refusal order.** Events are replayed in document order into the same checks,
so the first refusal inside a body is the one the reader route would raise.
Materialization stays after the parse. A decline replays nothing and the reader
then reads the body as before. Every reader error, limit and allocation
failure outside the body is raised by the same reader state as before. Only
allocation exhaustion can surface at a different point: the parser reserves an
admitted body's exact cell count once, before the body, instead of growing per
cell (the same typed `allocation("sparse worksheet cells", …)` error), and a
resumed reader's prefix-and-suffix copy is one new fallibly reserved buffer
with its own typed error (`worksheet lane resume`).

**Reduced readback.** For the admitted case every omitted span is a verbatim
source cell record inside the worksheet's own body, which compaction has just
recognized as lane-benign. The complete source parse (`Worksheet::store`, which
also validates styles) already accepted each such record with the same row
number, namespace bindings, style catalog and shared-string table. In the lane
grammar a record's parse depends only on its own bytes, its row number and, for
an inferred address, the preceding record's column; a replacement keeps that
column, the provenance writer emits explicit addresses for changed cells, and
rows whose membership changes are kept complete (change 0525). The reduced
document keeps the whole envelope, every row shell and every changed cell, so
the dimension, defaults, columns, merges, row ordering and each changed cell are
parsed and checked as before, and every recorded change is verified against the
reduced store. The two checks the complete parse makes on the whole document
are kept: a second record at an omitted address is refused by the collision
check, and the input-size limits cannot apply below the 16 MiB bound. ADR
0003's commit-time validation of the changed dependency closure is therefore
unchanged; what is no longer repeated is a second full parse of records that
were not written.

**Tests.** 46 new tests (33 in the original commits, 13 after review). `lane.rs`
pins the grammar, event stream, entry location and every
decline class. Differential suites run each pass twice, the second time with
the lane disabled, and compare complete `Store`/`Layout` `Debug`, compacted
bytes, web-proof eligibility and final bindings, or refusal `Debug` and
`Display`: 600 benign and 1,500 hostile generated worksheets for the parser and
scanner, 900 for compaction, namespace and envelope variants (strict dialect,
prefixed roots, foreign defaults, empty and duplicate `sheetData`, malformed
suffixes), event limits inside and after the body, invalid UTF-8,
oversized values, and ordering and coordinate refusals. The end-to-end suite
commits and saves random value, text, boolean, formula, clear, remove and style
edits on 70×64 and 80×60 sheets on both routes and requires identical published
bytes, reopened cells and refusals, including merge-follower refusals; it also
checks that inline-string, markup-compatibility, non-compact and formatted
sources keep the complete readback.

## Review fixes

An adversarial review — 33.1M mutated worksheets without a divergence between
the lane and the reader, 30M bodies matching `quick_xml` exactly, and 120,785
end-to-end commits identical — found the lane and its three replaying passes
sound and the reduced-readback admission weaker than its argument. Three commits
address every finding; each fix was reverted in turn and its new test failed
(`results/change-0744/post-review/mutation/mutation-results.json`).

1. **Admission on the worksheet's own body** (`3108e4dd40`). Compaction entered
   its lane at the first root child whose *local* name was `sheetData`, in any
   namespace. `<x:sheetData xmlns:x="urn:foreign"><row r="1"/></x:sheetData>`
   (or `<sheetData xmlns="urn:foreign">…`) placed before the real body therefore
   let a body with `<f>` formulas and `<is>` strings take the reduced readback.
   The lane now enters only where the parser and scanner do, and decides once
   there. Tests: the counterexample takes the complete parse; a benign body
   behind the same foreign element still takes the reduced readback; non-worksheet
   roots, nested, foreign-child and second-root cases decline.
2. **Whole-document checks** (`3108e4dd40`). (a) Markup-compatibility
   preprocessing in the complete parse refuses parts over 256 MiB. A
   268,435,452-byte compact source plus an 8-byte edit was refused on the base
   with `MarkupCompatibility(LimitExceeded("input bytes"))` but on the branch
   with the later web-extension limit error. The reduced route is now admitted
   only up to that web-extension reader's 16 MiB limit (`raw::web::MAX_XML_BYTES`),
   which every larger changed worksheet fails at the later step anyway, so the
   first error cannot move. Test: a part padded just past the limit never takes
   the reduced readback and fails identically on both routes (the 256 MiB
   counterexample itself is too large for the unit suite). (b) Change 0525's
   collision refusal is restored: a reduced store with a cell inside an omitted
   range is refused (`Store::avoids_omitted_cells`). Test: a fault-injected
   provenance writer copies the replaced cell's old record into the preceding
   omitted run and writes the new one after it; both routes report the complete
   parse's duplicate-cell refusal. With the check removed, the same faulty
   commit succeeded through the reduced readback.
3. **Entry location** (`3108e4dd40`). `lane::Entry::locate` now compares the
   start tag's name bytes, requires a delimiter after them, and declines any part
   beginning with a UTF-8 byte-order mark: `quick_xml`'s slice reader drops the
   mark without counting it, so its positions are three bytes short of document
   offsets there. All three lanes keep the reader for such parts. Tests: entry
   location, and byte-order-mark parity on every pass and end to end.
4. **Writer equivalence** (`b8367df8ca`). The end-to-end differential could not
   show that the provenance writer matches `rewrite`, because both of its routes
   use the provenance writer; a direct randomized differential now compares the
   two writers (see change 5 above).
5. **Packet.** The raw reports' `rustc 1.98.1` is explained under Measurements;
   `abba/run_abba.sh` now defaults to the `0744-before` binary the primary run
   used (it was always invoked with explicit `A` and `B`).
6. **Deterministic shared-formula groups** (`a755ff8c26`, pre-existing). The edit
   scanner ordered `Layout.shared_formulas` by `HashMap` iteration, which decided
   the reported group when several shared-formula refusals applied, and the
   layout's `Debug` form. Groups are now visited in document order of their first
   formula. Test: a two-group refusal names `si=7` on each of 33 scans.

## Measurements

Host AMD EPYC 9R45 (32 cores), Linux 7.0.0-1012-aws, CPU 12 pinned, other
agents active. Every binary was built by the repository's pinned Rust 1.95.0
(`rust-toolchain.toml`; each binary's `.comment` section reads `rustc version
1.95.0`). The raw reports' `rustc 1.98.1` is the ambient compiler outside the
repository, which the harness queries at run time from the process's working
directory; it did not build anything measured here. Harness unchanged.
Following the coordinator's
note on link-layout effects, the before leg is **not** the prebuilt base binary:
it is built with the identical command, features and target profile as the
after leg, from a detached base worktree whose path has the same length as the
branch worktree (446 embedded source paths). Binary SHA-256
(`results/change-0744/binaries.sha256`):

| role | SHA-256 |
| --- | --- |
| before, native (`targets/0744-before`) | `ac48ed6bde449374b76b9ed12f848ad9afeb403c2707ff1dda448b4229b36687` |
| after, native (`targets/0744`) | `5346833cba08a3d277d92da6eae49c15561cbf407154bfd167cbe2bed62eba13` |
| before, allocator metrics | `b90d82ec23d43782770aa193966fbbc22922201e27d93bd44dabfc65e47b722c` |
| after, allocator metrics | `cf41d655be0fc60fd5a6b44f7afbd22950a31734d57c61d023b7b074e6118ae6` |
| prebuilt base (superseded preliminary run only) | `fb535ebb4abb4154c4ac445b90a60399edf189d52cae76fea2947852a406cb0b` |

### Native timing

Eight processes per case in the order A1 B1 B2 A2 A3 B3 B4 A4; the four
adjacent pairs give the paired after/before p50 ratios, and the interval is a
percentile bootstrap (20,000 resamples, seed 744) of their mean. Samples within
a process are not treated as independent. Corpora `xlsx-dense-wide`
(2 × 256 × 256 numbers, 384,525 bytes, archive `5dd3ad70…`) and `xlsx-medium`
(4 × 32 × 32, 15,254 bytes, `9574867b…`); controls on the default cell-CRUD
`medium` corpus (`dfff7ec0…`).

| case | shape | samples/process | before median p50 | after median p50 | mean paired p50 change | bootstrap 95% | pair range |
| --- | --- | ---: | ---: | ---: | ---: | --- | --- |
| `xlsx_open_owned` | dense-wide | 200 | 1.340 ms | 1.348 ms | +1.32% | [−0.42%, +3.06%] | −0.45…+4.19% |
| `xlsx_first_cell` | dense-wide | 20 | 27.869 ms | 7.823 ms | **−71.91%** | [−72.20%, −71.62%] | −72.24…−71.48% |
| `xlsx_full_cell_scan` | dense-wide | 20 | 27.855 ms | 7.879 ms | **−71.81%** | [−72.21%, −71.38%] | −72.24…−71.11% |
| `xlsx_noop_commit_save` | dense-wide | 200 | 16.6 µs | 16.3 µs | +1.26% | [−3.13%, +5.65%] | −4.30…**+7.77%** |
| `xlsx_one_cell_commit_save` | dense-wide | 20 | 149.229 ms | 60.519 ms | **−59.53%** | [−59.67%, −59.38%] | −59.72…−59.31% |
| `xlsx_one_percent_commit_save` | dense-wide | 20 | 299.326 ms | 123.014 ms | **−59.03%** | [−59.22%, −58.90%] | −59.31…−58.86% |
| `xlsx_open_owned` | medium | 200 | 105.5 µs | 104.8 µs | −0.51% | [−0.73%, −0.21%] | −0.78…−0.08% |
| `xlsx_first_cell` | medium | 200 | 438.0 µs | 134.1 µs | **−69.49%** | [−69.72%, −69.10%] | −69.76…−68.92% |
| `xlsx_full_cell_scan` | medium | 200 | 439.3 µs | 134.9 µs | **−69.24%** | [−69.60%, −68.87%] | −69.75…−68.71% |
| `xlsx_noop_commit_save` | medium | 200 | 0.6 µs | 0.6 µs | +3.14% | [−3.38%, +11.40%] | −5.00…**+15.79%** |
| `xlsx_one_cell_commit_save` | medium | 200 | 2.146 ms | 870.2 µs | **−59.51%** | [−59.85%, −59.25%] | −60.02…−59.17% |
| `xlsx_one_percent_commit_save` | medium | 200 | 8.715 ms | 3.639 ms | **−58.09%** | [−58.44%, −57.62%] | −58.53…−57.40% |
| `xlsx_eager_cell_values_one_edit_save` (control) | cell-CRUD medium | 50 | 6.143 ms | 2.789 ms | **−52.85%** | [−55.43%, −48.67%] | −55.88…−46.56% |
| `xlsx_source_backed_cell_values_one_edit_save` (control) | cell-CRUD medium | 50 | 4.400 ms | 4.391 ms | −0.27% | [−0.80%, +0.24%] | −1.11…+0.36% |

The medium one-cell and one-percent cases stay below the handoff bound, so they
keep the complete readback; their gain is the lane alone. The source-backed
control, which plans through the shared traversal rather than the eager parser,
is unchanged.

**Regression flags over 5%.** Two no-op pairs exceed 5%: dense-wide pair 4
(+7.77%, 16.9 → 18.2 µs) and medium pair 1 (+15.79%, 570 → 660 ns). The no-op
commit returns before any worksheet work and its save raw-copies every member,
so no changed code runs in the timed region; both intervals include zero and
the same-arm spread is 5.4%/5.3% (before) and 13.7%/17.9% (after). A separate
confirmation ABBA with the same binaries and 2,000 samples per process
(`results/change-0744/abba/confirm/`) shows no regression:

| case | shape | samples/process | before median p50 | after median p50 | mean paired p50 change | bootstrap 95% | pair range |
| --- | --- | ---: | ---: | ---: | ---: | --- | --- |
| `xlsx_noop_commit_save` | dense-wide | 2,000 | 16.66 µs | 15.94 µs | −5.22% | [−8.80%, −2.22%] | −10.28…−1.33% |
| `xlsx_noop_commit_save` | medium | 2,000 | 0.57 µs | 0.56 µs | −1.26% | [−4.30%, +1.79%] | −5.08…+1.82% |
| `xlsx_eager_cell_values_one_edit_save` | cell-CRUD medium | 100 | 6.152 ms | 2.871 ms | −51.30% | [−54.79%, −44.79%] | −54.83…−41.47% |

The eager control's after arm has one slow process in both runs (B3 3.31 ms in
the primary run; same-arm spread 22%/21%); every pair still improves by at
least 41%. No other case has a pair above +5%; `xlsx_open_owned` dense-wide has
one pair at +4.19% with an interval spanning zero.

A preliminary run against the prebuilt base binary was stopped when the
layout note arrived and is retained, labelled superseded, in
`abba/prelim-prebuilt-base/`. On the eight processes it completed it gives
larger apparent gains (for example first-cell −72.92%, one-cell −60.16%)
because that binary is about 3% slower on untouched paths (first-cell before
median 28.80 ms against 27.87 ms); none of its numbers are used above.

### Post-review re-measurement

The review fixes add work to the measured commit path (the admission's element
checks and the collision check), so the dense-wide cases were measured again at
`a755ff8c26`. Both legs were rebuilt by the identical command, the before leg
again from an equal-length base worktree; builds of this workspace are not
bit-reproducible, so the rebuilt base digest differs from the first build's
(`results/change-0744/post-review/binaries.sha256`: before
`edb6f6ba4371bad7ab9410bd779fc85d9de4816c84eceecd503603073d0e93a4`, after
`7a718712e695be838a54691a1bad3297c50279343e482d76955d881898476f55`). Same design,
eight processes, 20 samples after three warm-ups:

| case | shape | samples/process | before median p50 | after median p50 | mean paired p50 change | bootstrap 95% | pair range |
| --- | --- | ---: | ---: | ---: | ---: | --- | --- |
| `xlsx_first_cell` | dense-wide | 20 | 27.928 ms | 7.892 ms | **−71.75%** | [−72.05%, −71.47%] | −72.20…−71.36% |
| `xlsx_full_cell_scan` | dense-wide | 20 | 28.271 ms | 8.071 ms | **−71.47%** | [−71.59%, −71.36%] | −71.64…−71.33% |
| `xlsx_one_cell_commit_save` | dense-wide | 20 | 151.133 ms | 60.688 ms | **−59.79%** | [−60.02%, −59.48%] | −60.08…−59.35% |
| `xlsx_one_percent_commit_save` | dense-wide | 20 | 303.553 ms | 123.169 ms | **−59.35%** | [−59.55%, −59.15%] | −59.60…−59.06% |

The paired changes match the primary run within 0.3 percentage points and no
pair is flagged. The allocator lane moves by exactly the collision check's one
range vector per worksheet verified through the reduced readback: dense-wide
one-cell 142,246 → 142,247 calls and 42,997,586 → 43,001,682 bytes, one-percent
303,182 → 303,184 calls and 89,541,943 → 89,570,935 bytes; medium, which keeps
the complete readback, and every region peak are unchanged
(`results/change-0744/post-review/alloc/alloc-summary.json`).

### Instruction counts

Callgrind over the timed region, one warm-up and two samples, matched
frame-pointer `profiling` builds (`results/change-0744/profile/`):

| region (dense-wide) | before Ir | after Ir | change |
| --- | ---: | ---: | ---: |
| one-cell commit+save, 3 operations | 7,205,638,296 | 2,358,903,627 | **−67.26%** |
| — `Edit::commit` | 5,866,976,585 | 1,020,237,162 | −82.61% |
| — `raw::worksheet::parse` (base and verification) | 2,958,012,099 | 552,045,449 | −81.34% |
| — edit layout scan | 1,556,466,723 | 359,787,877 | −76.88% |
| — compaction | 1,253,939,236 | 90,078,636 | −92.82% |
| — `PackageWriter::write_to_stream` | 1,338,661,069 | 1,338,665,823 | 0.00% |
| —— deflate | 773,088,367 | 773,095,551 | 0.00% |
| —— publication audit | 558,622,300 | 558,621,335 | 0.00% |
| first cell, 3 `Worksheet::store` calls | 2,472,250,164 | 914,820,256 | **−63.00%** |

Frame-pointer phase samples of the same one-cell region (30 samples each) move
from 20,634 to 8,459 inside the operation: scan 3,953 → 960, base parse
3,878 → 1,079, verification parse 3,819 → 23, compaction 2,458 → 146, while
deflate (5,047 → 5,075) and the audit (1,072 → 1,067) are unchanged. Deflate is
now 60.0% of the operation and the audit 12.6%; both belong to the OPC writer.

### Allocations

`litchi-perf-baseline-alloc`, operation region, five samples after two
warm-ups, processes A1 B1 B2 A2; every value is identical across samples and
processes (`results/change-0744/alloc/alloc-summary.json`):

| case | shape | allocation calls | allocated bytes | region peak live bytes |
| --- | --- | ---: | ---: | ---: |
| one-cell commit+save | dense-wide | 1,257,432 → 142,246 (−88.69%) | 141,889,103 → 42,997,586 (−69.70%) | 49,806,295 → 27,482,683 (−44.82%) |
| one-percent commit+save | dense-wide | 2,532,231 → 303,182 (−88.03%) | 286,855,356 → 89,541,943 (−68.78%) | 65,065,609 → 43,006,102 (−33.90%) |
| one-cell commit+save | medium | 21,384 → 5,072 (−76.28%) | 3,323,076 → 2,059,024 (−38.04%) | 1,266,339 → 1,266,332 (−0.00%) |
| one-percent commit+save | medium | 84,317 → 19,077 (−77.37%) | 13,041,962 → 7,999,258 (−38.67%) | 2,671,463 → 2,671,456 (−0.00%) |

These are allocator gauges, not RSS. Process RSS was not measured separately.

### Output identity

A standalone probe (`results/change-0744/output-identity/probe/`) replicates
the harness corpora and update coordinates and writes, for both shapes, the
archive, the first-cell view, a full stored-cell dump and the no-op, one-cell
and one-percent commit+save outputs with a full dump of each reopened output.
Built once against each leg's `litchi-xlsx`, all 18 artifacts are
byte-identical (`sha256-before.txt`, `sha256-after.txt`), and the archive
digests equal the harness corpus digests.

## What is not claimed

No claim is registered. The numbers are scoped to the named synthetic corpora,
this host and CPU 12. They do not establish behavior for:

* worksheets outside the lane subset: formulas, inline strings (which this
  library's writer produces for text edits, so a sheet with edited text keeps
  the reader route afterwards), comments, CDATA, entities or unusual markup take
  the unchanged reader route at unchanged cost;
* Excel-produced sheets that mention the markup-compatibility namespace or carry
  `x14ac:dyDescent`: their bodies take the lane, but the pre-existing
  `process_ooxml` rewrite and x14ac capture passes remain and were not measured;
  their first commit also keeps the complete readback because compaction
  rewrites the declaration line break;
* RSS, cold cache, other producers, source-backed reads, concurrency or other
  platforms.

The publication audit and deflate are unchanged and now dominate the commit
path; they are owned by the OPC writer (change 0747 is concurrently working on
source-backed publication audits; ADR 0031 covers parallel deflate).

## Follow-up opportunities

* Admit plain inline strings (`<is><t>…</t></is>`, with `xml:space="preserve"`)
  in the lane, so sheets this library has written text into stay on it.
* Build only touched rows' cell slots in the edit scanner: the writer copies
  untouched rows as whole spans, so their per-cell slots serve only the dimension
  bound (about 11% of the remaining one-cell operation).
* Replay the parser lane in one pass with rollback instead of `recognize` then
  `walk` (about 13% of a parse), and shrink the 184-byte `Stored` record.
* The lane for the x14ac capture and MCE preprocessing passes that Excel files
  still take.
* Pre-existing, found in review: eager cell edits on worksheets that begin with a
  UTF-8 byte-order mark fail on both routes, because the edit scanner records
  `quick_xml` reader positions, which exclude the mark, as byte offsets; its
  spans are three bytes short. The lanes decline such parts, so the behavior is
  unchanged by this change.

## Authority

ADR 0003 (commit validates the changed dependency closure; no mutation of the
source snapshot), ADR 0005 (measurement contract; no new cache or ambient
behavior), ADR 0006 (preservation by default; validation never mutates; typed
refusals), ADR 0008 (fresh gates), ADR 0011 and 0024 (no archive types cross
the facade; no new dependency edge). Owner decisions of change
[0652](0652-owner-decisions-for-the-third-wave.md): standing trade-off 2
(correctness first — every refusal, limit and output byte is kept, and every
fast route falls back to the reader) and trade-off 3 (optimize the benign path;
anything outside it pays exactly what it paid before). The reduced readback
reuses change 0525's accepted mechanism and change 0025's handoff bounds.

## Verification

Thirteen gates pass at `a755ff8c26`, run after the review fixes from an empty
target directory (`results/change-0744/gates.txt`): `cargo fmt --all --check`;
`cargo check` of `litchi-xlsx` (all targets, default and all features) and of
the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`; warning-denied Clippy on
the library and all targets; `cargo test -p litchi-xlsx` (1,414 passed), with
all features (1,433 passed) and the facade (382 passed, 7 ignored);
warning-denied rustdoc; crate boundaries (64 packages, 241 declarations, 11
explicit debts); the non-iWork gate; and the structural claims check (10
claims). The same gates passed at `6ac6b92d8a` before review (1,401, 1,420 and
382 tests). The harness is unchanged, so its own tests and the coverage
validator were not rerun.

## Cleanup

Binary digests are recorded above. After each round — the original change and
the review fixes — the target directories (`targets/0744`, including its
`gates` directory, `targets/0744-before`, the two probe targets), the detached
before worktree and the scratch directory, including every `perf.data`,
callgrind output and corpus copy, are removed after the evidence was copied;
the branch worktree is kept (`results/change-0744/cleanup.json`).
