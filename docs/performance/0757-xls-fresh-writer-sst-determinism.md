# 0757: The fresh XLS writer lists shared strings in first-occurrence order and refuses, when they are given, strings its shared-string, formula, number-format, sheet-name and defined-name fields cannot hold: the same workbook is written byte for byte in every process, those strings are never cut, and string-heavy writes execute 11–15% more instructions

Status: retained, implemented. `performance_claim: none` — this is a
correctness change (ADR 0006 determinism; GOAL.md rule 3, "never trade a typed
refusal for a partial or guessed edit"). Its costs are measured below as
evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `9ff78bbf1c` (the head of `perf/0753-legacy-fresh-writer-text-paths`);
branch `perf/0757-xls-fresh-writer-sst-determinism`. Commits:

| commit | what it does |
| --- | --- |
| `15e50e3184` | `fix(xls)`: shared strings in first-occurrence order; `Error::StringTooLong`; typed refusals in every string encoder that cut or wrapped |
| `7309029a6d` | `fix(doc,ppt,xls)`: the 0753 review's follow-ups (section "0753 review follow-ups") |
| `921787e2c0` | `perf(xls)`: the order's sort keyed by one packed integer (same order, same bytes) |
| `a400341bdf` | `fix(xls)`: data-validation strings with Latin-1 characters written as Latin-1 bytes |
| `8a2693a83e` | `test(xls)`: defined-name records and `NamePublish` names at N − 1, N and N + 1 |
| `6cebd664e6` | `fix(xls)`: after the 0757 review, number formats, cell styles and defined names are refused when registered, not when written (section "Refused at registration") |

Every timing and counter below is of `921787e2c0` against an identically built
base. `a400341bdf` only changes the data-validation string encoder, which no
measured case reaches; the harness at `a400341bdf` executes 17,060,215 user
instructions per `xls_fresh_write_to/large` iteration against 17,033,997 for the
measured build (+0.15%, `gates.txt`). `6cebd664e6` (after the review) changes
only registration calls none of the measured cases make, plus one `is_empty`
test per `Format` record written (eight per measured workbook); it was not
re-measured. Evidence:
[results/change-0757](results/change-0757/README.md).

## Result

**Determinism.** Record 0753 found that the fresh XLS writer listed each
worksheet's strings in its cell map's iteration order, which depends on the
map's per-process random seed. It does. Eight processes of the base writing the
same four-worksheet workbook of 80,000 strings wrote eight different byte
streams; eight processes of this branch write one (probe, `distinct outputs`
columns below). The multi-string golden test added here fails on the base with
"many_strings is not deterministic" (`behaviour/base-goldens-final.txt`) and
passes in twelve separate processes on the branch, each with fresh hash seeds.

**No string is cut in the fields this record owns.** Each string field in the
table below now either fits or is refused with the typed `Error::StringTooLong
{ field, utf16_units, limit }` (or `InvalidData` for an empty number format)
before anything is written. Other string writers still cut, miscount or
misencode (font names, AutoFilter strings, internal hyperlinks, PivotTable
names; section "Findings outside this change"). What the base did with the
same inputs (`behaviour/base-probe.txt`, the probe
`probe-src/string_fields_probe.rs`):

| field (record) | BIFF8 limit | the base, one unit or more past it | this branch |
| --- | --- | --- | --- |
| shared string (`SST`, `XLUnicodeRichExtendedString.cch`) | 65,535 UTF-16 units | 65,536 × `b` read back as 65,535; 70,000 × `é` as 65,535; 65,534 × `a` + one emoji cut through the surrogate pair, and litchi's reader then refused the **whole workbook** ("lone surrogate found") | refused at `write_string`, and for any other string cell while staging the table |
| formula string constant (`PtgStr`) | 255 | 300 × `a` written as 255; 254 × `a` + emoji as the 254 `a`s | refused (`FormulaTokenizer::tokenize`, `encode_ptg_tokens`) |
| number format (`Format.stFormat`) | 1–255 | 256 units written, then refused by litchi's reader on open; 70,000 units panicked on a `u16` overflow (debug build; wraps in release); an empty format was written and refused on open | refused by `register_number_format` / `add_cell_style` (since `6cebd664e6`; the writer stays able to write), and again by the encoder |
| defined name (`Lbl`) | 255 | `Café` and `Rocket😀Launch` written with the wrong length, reader refused the workbook; 200 emoji (400 units) accepted and corrupt | written correctly; 400 units refused |
| worksheet name (`BoundSheet8`) | 31 | 16 × `é` (32 bytes, 16 units) refused as "1-31 characters"; the encoder sliced names at byte 31 | 31 units accepted; 32 refused |

Comment text and author, headers and footers, the defined-name comment and the
`NameFnGrp12`/`NamePublish` names already refused over-long input; they keep
those refusals, now tested at N − 1, N and N + 1, and their encoders use
checked lengths instead of wrapping casts.

**Bytes.** Unchanged for every worksheet with at most one distinct string: the
pre-0753 golden fixtures still match (one fixture's two over-limit strings are
now refused and were replaced by strings exactly at the limit, whose digest was
recorded from the base, `9ff78bbf1c`), and every harness corpus is identical in
both legs (`same output` column). Changed, by design: worksheets with two or
more distinct strings (SST order and `LabelSst` indices, now reproducible);
defined names with Latin-1 or supplementary-plane characters and
data-validation strings with Latin-1 characters (previously unreadable).

**Cost.** Worksheets with many strings pay one extra sort of their string
cells: +11.4% to +15.2% exact instructions per `write_to` on 80,000 string cells,
0.4–12.2% wall time depending on the process's heap state (below). The
harness's numeric and single-string XLS shapes stay within 1.6% of the base in
wall time (the probe's `write_to`-only tiny case: 3.7%, +5.5% instructions),
except `xls_fresh_write_to/large`, where a glibc heap-trim effect (not more
work: 3.4% fewer instructions) costs 5–8%; with glibc's trim and mmap
thresholds pinned the branch is 1.7% faster there.

## Why this record exists

Record 0753 (hash-once shared strings) listed two pre-existing defects of the
fresh XLS writer as findings outside its change: the process-dependent SST
order, which conflicts with ADR 0006 ("Serialization is deterministic unless a
`Clock`, actor identity, or cryptographic RNG is explicitly supplied"), and the
silent truncation of SST strings at 0xFFFF UTF-16 units. The coordinator
dispatched this record to fix both and to audit the writer's other string
encoders for the same fault. The 0753 review then found that the truncation is
worse than 0753 said (the lone-surrogate case above) and asked for four
follow-ups, folded in here as a separate commit.

## What changed

Scope: `crates/litchi-xls` (writer, error type, and one line of the workbook
editor), and for the 0753 follow-ups `crates/litchi-doc` and `crates/litchi-ppt`
writers. No `unsafe`, no dependency, no limit relaxed, no ambient behaviour.

### Shared-string order (`writer/core/stream/shared_strings.rs`, `codec.rs`)

- `SharedStringTable::build` collects each worksheet's string cells, sorts them
  by `row << 16 | column` (exact for every `u32` row and `u16` column, and
  ordered as the pairs are), and assigns each distinct string its index at its
  first occurrence: worksheets in order, rows, then columns — the order the cell
  records are written in. The hash map stays keyed by the same randomly seeded
  SipHash-1-3 (hashing each string cell once, as 0753 did); its seed no longer
  reaches the output.
- The table records each string cell's index in that order; the emission loop,
  which already sorts cells by row and column, counts string cells as it meets
  them (before any branch can skip one, so a string written over a data-table
  anchor keeps the later ordinals aligned) and takes the recorded index. As in
  0753, `index_for` checks the recorded index by content and falls back to the
  hashed lookup, so the output never depends on the two orders agreeing. The
  emission tuples drop 0753's ordinal (16 instead of 24 bytes per cell).
- `build` returns `Result`: a distinct string longer than 65,535 UTF-16 units is
  refused at its first occurrence, before any byte of the workbook stream
  exists. This covers string cells that do not come through `write_string`
  (pivot labels, table headers, or a crate-internal insertion).

### Typed refusals (`writer/string_limits.rs`, `error.rs`, the encoders)

- `Error::StringTooLong { field: &'static str, utf16_units: usize, limit:
  usize }`, displayed as "`{field}` has `{n}` UTF-16 code units; BIFF8 stores at
  most `{limit}`", is the refusal every new or recounted length check returns.
- `writer/string_limits.rs` holds the limits with their MS-XLS sections and
  three helpers: `checked_utf16_len`, `ensure_utf16_len_within` (which counts
  code units only when the UTF-8 length alone cannot prove the string fits: no
  UTF-8 byte yields more than one code unit), and `u8_len`/`u16_len` narrowing
  conversions that refuse instead of wrapping.
- Shared strings: `write_string` and `write_string_with_format` refuse before
  touching the cell (a refused write leaves the previous value);
  `write_sst` validates every string before writing any byte, and no longer has
  a `.min(0xFFFF)`/`take(0xFFFF)` path.
- Formula strings: `FormulaTokenizer::tokenize` refuses a string constant over
  255 units; `encode_ptg_tokens` now returns `Result<Vec<u8>, Error>` and
  refuses one too (it no longer truncates and drops a split surrogate); the
  array-formula encoder narrows its count with `u8_len`. The conditional-format
  path, which wraps tokenizer errors in `InvalidData` with context, passes
  `StringTooLong` through unchanged.
- Number formats: `write_format_record` refuses more than 255 units (the limit
  of MS-XLS 2.4.126 and of the crate's reader) and computes the record length
  without the overflowing `u16` arithmetic.
- Worksheet names: `add_worksheet` counts UTF-16 units, 1 through 31 (MS-XLS
  2.4.28, and the rule the workbook editor already applies); `write_boundsheet`
  refuses instead of slicing at byte 31.
- Defined names: `define_name*` count UTF-16 units (255 at most); `write_name`
  sets `cch` from the UTF-16 count and writes a compressed name as the low byte
  of each code unit (Latin-1) instead of its UTF-8 bytes;
  `write_defined_name_record`, `write_name_function_group` and
  `write_name_publish` narrow their counts with checks.
- Comments: `write_txo` and `write_note` narrow their counts with checks (the
  insertion refusals, 65,535 and 54 units, are unchanged).
- Data validation (`a400341bdf`): `write_unicode_string_biff8` wrote the UTF-8
  bytes of a compressed string; a prompt titled `Café` made litchi's reader drop
  the worksheet ("Worksheet 'Sheet index 0' not found",
  `behaviour/findings-probe.txt`). It now writes Latin-1 bytes.
- The workbook editor's `insert_formula` shares the formula encoder, so it too
  refuses a string constant over 255 units instead of staging a truncated one
  (`cell_values/tests.rs`).

### Public API changes (owner decision 1: breaking changes are acceptable)

- New variant `litchi_xls::Error::StringTooLong`.
- `writer::formula::encode_ptg_tokens` returns `Result<Vec<u8>, Error>`
  (all callers are in `litchi-xls`).
- `FormulaTokenizer::tokenize`, `Writer::write_string`,
  `write_string_with_format`, `define_name`, `define_name_local`,
  `define_name_with_comment`, `add_worksheet`, `write_to` and `save` refuse
  strings they used to cut or corrupt; `add_worksheet` now accepts 16–31-unit
  non-ASCII names it used to refuse, and its and `define_name*`'s length
  refusals are `StringTooLong` instead of `InvalidData`.
- Since `6cebd664e6`: `Writer::register_number_format`, `Writer::add_cell_style`,
  `FormattingManager::register_number_format` and
  `FormattingManager::register_cell_style` return `Result<u16>`, and
  `define_name*` refuse an unsupported reference or an over-long comment when
  called instead of at write time.

### The threshold for cell text

BIFF8 stores an SST entry's length in a 16-bit `cch` (MS-XLS 2.5.293), and MS-XLS
sets no smaller bound on SST entries. The crate reads SST entries of any 16-bit
length, its workbook editor refuses a new shared string over `u16::MAX`
("shared string exceeds u16 characters"), and the fresh writer's comment and
shape text use 65,535 too; no cell-text path in `litchi-xls` enforces anything
stricter. The fresh writer therefore refuses at 65,535 UTF-16 units. Excel's
documented application limit of 32,767 characters per cell is not a format
limit and is not enforced by any `litchi-xls` path; MS-XLS's own 32,767 bounds
apply to records this writer does not emit (the formula-result `String` record,
phonetic `ExtRst` text, some PivotTable strings). Whether Excel opens a
40,000-character cell is not verified here; litchi writes and reads it whole.

### Refused at registration (`6cebd664e6`, after the 0757 review)

The review found that the first version refused a string only when the
workbook was written, which left the writer permanently unable to write:
`register_number_format` and `add_cell_style` could not fail, no API removes a
format, and after `register_number_format(&"0".repeat(256))` every later
`write_to` failed; `define_name_with_comment`'s comment was likewise refused
only by the `NameCmt` encoder. Now:

- A number format must hold 1 through 255 UTF-16 code units when it is
  registered (`formatting::validate_number_format`): an empty pattern is
  `InvalidData`, a longer one `StringTooLong`. This also stops the writer
  emitting an empty custom format, which litchi's reader refused (the finding
  the first version recorded). `register_cell_style` registers the style's
  number format first — the only fallible step; the font and XF tables are
  separate and keep their indices — so a refused style adds no font, format or
  XF.
- `define_name`, `define_name_local` and `define_name_with_comment` store a name
  only after its reference encodes (`DefinedName::to_biff_formula`, the call the
  write makes) and its comment fits `NameCmt` (255 units, MS-XLS 2.4.176).
  `remove_name` could already take a bad name back out; now none gets in.
- `write_format_record` keeps its own length check and also refuses an empty
  string, as defence in depth behind registration.
- Tests: refused registrations (empty and 256-unit formats in four encodings,
  cell styles carrying them, an over-long comment, four unsupported references
  through all three `define_name*` calls) leave the writer writing, byte for
  byte, what a writer that never saw the calls writes; the manager's format,
  font and XF tables are unchanged after each refusal; the encoder refuses what
  registration refuses when a format is placed behind it.

## 0753 review follow-ups (`7309029a6d`)

1. **Severity of the truncation.** 0753's finding is amended in place with the
   lone-surrogate consequence and a pointer here; the regression test
   `a_string_whose_surrogate_pair_straddles_the_limit_is_refused_not_split`
   reproduces the review's case (65,534 × `a` + one emoji is refused; 65,533 ×
   `a` + one emoji is written whole and read back).
2. **Reservations.** DOC reserved the `WordDocument` stream before
   `build_revision_writer_data()`; it now reserves after that document-wide
   check. Both DOC (`TextStream::try_new`) and PPT (`reserved_document_stream`)
   reserve through `try_reserve_exact`, capped at 4 GiB (the reach of their
   32-bit offsets; a larger estimate only means growth, as before), so a
   reservation that cannot be made is `WriteError::InvalidData("… allocation is
   too large")` instead of an abort. Story-level checks still run as each story
   is appended — hoisting them would need a second pass over the text — so a
   document refused by a story check still makes its reservation first, now
   fallibly. The capacities are unchanged, so is every byte (the DOC and PPT
   pre-0753 goldens pass).
3. **`InPlaceRecord`.** `begin` and `finish` are private; the only constructor
   is `InPlaceRecord::write(output, version, instance, type, |output| body)`,
   which patches the header whenever the body succeeds, so a forgotten `finish`
   no longer compiles; the type is `#[must_use]`. The slide, `PPDrawing`,
   `DgContainer`, root `SpgrContainer`, `SpContainer` and text-atom call sites
   use it (the long bodies moved into `append_shape_container_children` and
   `append_dg_container_children` unchanged). A debug assertion in `Drop` was
   not used: records are legitimately abandoned on every error path, where the
   output is discarded.
4. **Direct tests.** `utf16_units` and `contains_field_character` against naive
   references over all 1,112,064 Unicode scalars (4,382,592 UTF-8 bytes) at
   every offset from a 64-byte boundary, plus every split of each UTF-8 length
   class across the real 127-byte count blocks and each field character at every
   offset through the second 256-byte scan block; `InPlaceRecord` against
   `RecordBuilder` bytes, nested, and on a failed body; both reservation
   refusals. The multi-string XLS goldens asked for are
   `multi_string_worksheets_are_pinned_across_processes` (0757's commit).

## Why it is sound

- **Order is a pure function of the cells.** The sort key is injective on
  `(row, column)` and totally ordered, the traversal is worksheet order, and
  distinct keys make the unstable sort exact. Tests: an explicit expected order
  for cells inserted in no particular order; a reference model that sorts the
  keys and indexes each string at its first occurrence, over 319 string cells
  with repeated, empty, Latin-1, CJK, supplementary and CONTINUE-spanning
  strings; the same cells in eight writers (eight map seeds) give identical
  tables; two pinned golden digests of multi-string workbooks (3,600 and 450
  string cells), each also rebuilt four times in-process; the reader's SST
  equals the first-occurrence list.
- **Every cell still reads back.** 3,600 cells of four worksheets and the
  scrambled fixture are read back cell by cell; a string written over a
  data-table anchor leaves the cells after it correct; the recorded-index path
  is hit (callgrind: one `hash_one::<&str>` per string cell per write, 320,000
  calls over four writes of 80,000 cells).
- **Refusals precede output and never mutate.** API refusals happen before the
  cell, worksheet, name, format or style is stored (tests check the previous
  value, the collection sizes, and since `6cebd664e6` that the writer still
  writes the bytes of a writer that never saw the refused calls); save-time
  refusals, now only for strings that bypass the API (such as pivot labels),
  leave the destination empty (tests check `Cursor` length 0); `write_sst`
  validates all strings before writing.
  Limits are tested at N − 1, N and N + 1 for ASCII, Latin-1, CJK and a
  surrogate pair at the end, including the pair that straddles the limit, with
  the written strings read back whole.
- **Hashing is unchanged**: the same keyed SipHash-1-3 as 0753; no new hasher.
- **No other hash order reaches the output.** The remaining `HashMap`/`HashSet`
  iterations in the writer feed `BTreeSet`s, min/max folds or membership tests
  (checked by reading every iteration over `cells`, `column_widths`,
  `hidden_columns`, `row_heights` and `hidden_rows`).

## Authority and constraints

- ADR 0006: deterministic serialization; normal save never repairs (the writer
  refuses rather than truncating).
- GOAL.md rule 3 ("never trade a typed refusal for a partial or guessed edit")
  and rule 12 (typed errors, no weakened limits): limits are enforced, none
  relaxed; the worksheet-name change counts the unit MS-XLS counts.
- ADR 0016: invalid inputs are refused without changing the writer's state.
- ADR 0012: formula references stay valid by construction; `encode_ptg_tokens`
  becomes fallible only for string constants.
- Change 0652: trade-off 1 (the breaking API changes above), 2 (correctness
  first: string-heavy writes pay for determinism) and 3 (the common path keeps
  its cost: numeric and single-string sheets are within noise in instructions).

## Measured

Harness `tools/perf-baseline` built at each leg with one command
(`scripts/build_leg.sh`: `cargo build --release --locked --offline
--manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline`,
`CARGO_BUILD_JOBS=6`), before from a detached worktree at `9ff78bbf1c`, after at
`921787e2c0`. SHA-256 `d49cc06b…f6a4a` (before) and `e276dfcb…927230` (after);
probes `5f53619e…c5371` and `082b3bb9…a53` (`binaries.sha256`). Binaries at
equal-length paths (`bin/A/lpb`, `bin/B/lpb`), every process pinned to CPU 16.

### Paired wall time

ABBA (A B B A A B B A) per case, median of per-process p50s, median of the four
adjacent-pair ratios and a percentile bootstrap. Window 1 12:34–12:35Z, window 2
12:42–12:43Z, host shared (load average 1.3–3.5).

| selector | shape | warm-up + samples | before p50 ms | after p50 ms | window 1 ratio [95% CI] | window 2 ratio [95% CI] | same output |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| `xls_fresh_write_to` | tiny | 100+1000 | 0.0049 | 0.0050 | 1.015 [1.012, 1.018] | 1.016 [1.000, 1.023] | yes |
| `xls_fresh_write_to` | large | 10+100 | 1.0986 | 1.1843 | **1.079** [1.074, 1.083] | 1.053 [0.970, 1.074] | yes |
| `xls_fresh_write_to` | payload-heavy | 5+40 | 2.3773 | 2.4078 | 1.013 [1.000, 1.018] | 1.012 [1.010, 1.014] | yes |
| `xls_semantic_one_edit_save` (control) | tiny | 40+400 | 0.0496 | 0.0496 | 0.997 [0.995, 1.003] | 0.995 [0.993, 0.998] | yes |
| `xls_semantic_one_edit_save` (control) | large | 10+100 | 1.5287 | 1.5164 | 0.991 [0.987, 0.998] | 0.995 [0.985, 0.997] | yes |
| `doc_fresh_write_to` | tiny | 100+1000 | 0.0058 | 0.0058 | 0.994 [0.981, 1.003] | 0.987 [0.963, 0.998] | yes |
| `doc_fresh_write_to` | large | 20+200 | 0.1425 | 0.1362 | 0.956 [0.954, 0.958] | 0.964 [0.960, 0.972] | yes |
| `doc_fresh_write_to` | payload-heavy | 5+40 | 1.6117 | 1.6123 | 1.000 [0.995, 1.008] | 1.001 [0.998, 1.006] | yes |
| `ppt_fresh_write_to` | tiny | 100+1000 | 0.0098 | 0.0095 | 0.972 [0.940, 1.025] | 0.942 [0.924, 0.999] | yes |
| `ppt_fresh_write_to` | large | 20+200 | 0.0717 | 0.0738 | 1.023 [1.015, 1.087] | 1.024 [1.014, 1.043] | yes |
| `ppt_fresh_write_to` | payload-heavy | 5+40 | 0.8060 | 0.6374 | **0.792** [0.775, 0.801] | 0.805 [0.784, 0.810] | yes |

The harness's timed region builds the writer and writes it. The DOC and PPT
rows measure only the follow-ups' structural changes; section "Regression
flags" explains the two bold rows.

### Hardware counters per iteration

`perf stat` at 20/220 samples (200/2,200 for tiny, 40/440 for the tiny
control), no warm-up, A B B A, differenced and averaged (`counters/`). An
iteration includes the harness's untimed work.

| selector | shape | user instructions | user cycles | kernel instructions | page faults |
| --- | --- | --- | --- | --- | --- |
| `xls_fresh_write_to` | tiny | 82,289 → 84,833 (+3.1%) | 24,296 → 26,001 (+7.0%) | 33,183 → 32,950 (−0.7%) | 0 → 0 |
| `xls_fresh_write_to` | large | 17,649,538 → 17,057,174 (−3.4%) | 4,556,245 → 4,740,057 (+4.0%) | 1,034,411 → 2,347,696 (+127.0%) | 188 → 435 |
| `xls_fresh_write_to` | payload-heavy | 21,374,377 → 21,390,839 (+0.1%) | 7,648,428 → 7,646,418 (−0.0%) | 11,083,450 → 11,158,970 (+0.7%) | 2,103 → 2,119 |
| `xls_semantic_one_edit_save` | tiny | 1,534,817 → 1,526,230 (−0.6%) | 517,183 → 503,027 (−2.7%) | 38,639 → 34,572 (−10.5%) | 0 → 0 |
| `xls_semantic_one_edit_save` | large | 163,810,721 → 162,374,080 (−0.9%) | 46,323,702 → 46,460,324 (+0.3%) | 2,889,032 → 2,641,112 (−8.6%) | 516 → 468 |
| `doc_fresh_write_to` | tiny | 96,337 → 96,036 (−0.3%) | 31,142 → 24,642 (−20.9%) | 33,984 → 33,532 (−1.3%) | 0 → 0 |
| `doc_fresh_write_to` | large | 3,091,954 → 3,005,859 (−2.8%) | 668,747 → 639,218 (−4.4%) | 40,062 → 29,942 (−25.3%) | 0 → 0 |
| `doc_fresh_write_to` | payload-heavy | 8,392,933 → 8,390,049 (−0.0%) | 4,675,016 → 4,642,348 (−0.7%) | 13,291,416 → 13,244,546 (−0.4%) | 2,517 → 2,517 |
| `ppt_fresh_write_to` | tiny | 178,868 → 178,192 (−0.4%) | 41,850 → 41,597 (−0.6%) | 35,350 → 33,304 (−5.8%) | 0 → 0 |
| `ppt_fresh_write_to` | large | 1,365,533 → 1,373,150 (+0.6%) | 313,807 → 284,270 (−9.4%) | 66,902 → 76,544 (+14.4%) | 2 → 2 |
| `ppt_fresh_write_to` | payload-heavy | 7,285,594 → 7,085,564 (−2.7%) | 3,704,551 → 3,446,720 (−7.0%) | 638,885 → 598,213 (−6.4%) | 113 → 106 |

### `write_to` alone: probe timings, exact instructions, allocations

The harness corpora cannot show the cost of the new order: their worksheets
hold no string or one. The probe (`probe-src/main.rs`, built against each leg
with one command, `scripts/build_probe.sh`) builds a writer once and times only
`write_to`; it adds `xls_multi_string/distinct` (four worksheets of 2,500 × 8
string cells, all 80,000 distinct) and `xls_multi_string/repeated` (the same
grid drawn from 64 labels). Instructions are callgrind's, differenced over two
iteration counts (`probe/callgrind/`); they vary ±1.5% between runs because the
maps' random seeds change their probe sequences.

| probe case | processes | before p50 ms | after p50 ms | paired ratio [95% CI] | instructions per write | distinct outputs before / after |
| --- | ---: | ---: | ---: | ---: | ---: | --- |
| `xls_multi_string/distinct` | 16 | 12.902 | 14.470 | **1.122** [1.115, 1.126] | 166,172,723 → 188,615,049 (+13.5%) | 8 / 1 |
| same, malloc thresholds pinned | 16 | 12.759 | 12.815 | 1.004 [1.000, 1.012] | 165,148,406 → 187,023,776 (+13.2%) | 8 / 1 |
| `xls_multi_string/repeated` | 16 | 10.426 | 10.994 | **1.057** [1.035, 1.081] | 141,101,098 → 160,299,156 (+13.6%) | 8 / 1 |
| same, malloc thresholds pinned | 16 | 10.255 | 10.867 | **1.061** [1.055, 1.064] | 141,085,478 → 162,531,220 (+15.2%) | 8 / 1 |
| `xls_fresh_write_to/tiny` | 8 | 0.0044 | 0.0046 | 1.037 [1.031, 1.045] | 60,519 → 63,851 (+5.5%) | 1 / 1 |
| `xls_fresh_write_to/large` | 8 | 0.5618 | 0.5391 | 0.961 [0.954, 0.965] | 9,884,364 → 9,132,825 (−7.6%) | 1 / 1 |
| `xls_fresh_write_to/payload-heavy` | 8 | 0.9548 | 0.9524 | 0.997 [0.996, 1.000] | 22,694,088 → 22,699,619 (+0.0%) | 1 / 1 |
| `doc_fresh_write_to/tiny` | 8 | 0.0069 | 0.0068 | 0.991 [0.983, 1.016] | 102,733 → 103,232 (+0.5%) | 1 / 1 |
| `doc_fresh_write_to/large` | 8 | 0.0954 | 0.0952 | 0.996 [0.989, 1.001] | 1,742,653 → 1,735,214 (−0.4%) | 1 / 1 |
| `doc_fresh_write_to/payload-heavy` | 8 | 0.3359 | 0.3341 | 0.995 [0.994, 0.997] | 6,536,234 → 6,534,306 (−0.0%) | 1 / 1 |
| `ppt_fresh_write_to/tiny` | 8 | 0.0148 | 0.0148 | 1.000 [0.997, 1.001] | 170,597 → 170,804 (+0.1%) | 1 / 1 |
| `ppt_fresh_write_to/large` | 8 | 0.0957 | 0.0957 | 1.000 [0.996, 1.002] | 1,034,958 → 1,026,034 (−0.9%) | 1 / 1 |
| `ppt_fresh_write_to/payload-heavy` | 8 | 0.4081 | 0.4067 | 0.997 [0.987, 1.000] | 14,042,781 → 14,060,829 (+0.1%) | 1 / 1 |

The eight-process run of the two multi-string cases (`probe/summary.json`) gave
1.124 [1.098, 1.138] and 1.039 [1.032, 1.061]; a development run with the base's
per-process p50s spread from 13.3 to 20.1 ms is not retained. Attribution
(callgrind, distinct, per write): the new sort of string cells is 20.0M
instructions, the cheaper 16-byte emission tuples save 9.8M, and string hashing
now compiles out of line in the table build (+≈8M; an `#[inline]` hint on
`SharedStringTable::key` did not change it and is not retained). Packing the
sort key (`921787e2c0`) had brought the added cost from 29.8M to 19.6M
instructions per write for `distinct` and from 31.0M to 20.5M for `repeated`
(development callgrind runs of an XLS-only probe, not retained).

Counting global allocator, per write (`probe/summary.json`): allocations are
unchanged on every case except +1 on XLS payload-heavy (the string-cell scratch
vector) and +5/+4 on the multi-string cases; allocated bytes fall 65,541 on XLS
large (65,536 of them four worksheets × 2,048 cells × 8 bytes of tuple width);
peak live bytes
are unchanged everywhere except −364,288 (−2.8%) on `distinct`. DOC and PPT
allocations are identical.

### Heap-state check

`scripts/tunables_check.py`: the same harness processes with glibc's default
thresholds and with `trim_threshold` and `mmap_threshold` pinned at 256 MiB
(`tunables/`), A B B A:

| case | thresholds | timed p50 ms, A B B A | page faults per iteration | user instructions per iteration |
| --- | --- | --- | --- | --- |
| `xls_fresh_write_to/large` | default | 1.097 / 1.180 / 1.189 / 1.113 | 208 / 414 / 435 / 230 | 17.69M / 17.05M / 17.08M / 17.65M |
| `xls_fresh_write_to/large` | pinned | 0.991 / 0.978 / 0.976 / 0.997 | 0 / 0 / 0 / 0 | 17.63M / 17.06M / 17.03M / 17.68M |
| `ppt_fresh_write_to/payload-heavy` | default | 0.687 / 0.655 / 0.685 / 0.668 | 113 / 106 / 106 / 113 | 7.29M / 7.09M / 7.09M / 7.29M |
| `ppt_fresh_write_to/payload-heavy` | pinned | 0.631 / 0.686 / 0.672 / 0.666 | 6 / 0 / 0 / 6 | 7.10M / 7.35M / 7.35M / 7.10M |

## Regression flags (every case above 5%, and every smaller measured increase)

- **`xls_fresh_write_to` large: 1.079 and 1.053 wall time**, +4.0% user cycles,
  page faults 188 → 435 and kernel instructions +127% per iteration, with 3.4%
  *fewer* user instructions. With glibc's trim and mmap thresholds pinned, both
  legs take no page faults and the branch runs 0.983 of the base; the probe,
  which times `write_to` alone, measures 0.961 and −7.6% instructions. The
  emission's sorted cell list is now 16 instead of 24 bytes per cell (65,536
  bytes less per write here), which changes where glibc trims the heap between
  iterations of this harness loop. Reported, not attributed to more work.
- **`xls_multi_string/distinct` and `/repeated` (probe): +11.4% to +15.2%
  instructions per write; wall 1.004–1.122 and 1.039–1.061** depending on heap
  state. This is the price of determinism for string-heavy worksheets: one sort
  of each worksheet's string cells. Not covered by any harness selector.
- **`xls_fresh_write_to` tiny: +5.5% instructions per write (probe; +3.1% per
  harness iteration), wall 1.015/1.016 (harness) and 1.037 (probe).** About 425
  instructions are the nine new length checks (eight number formats and the
  sheet name); most of the rest is `Vec::extend_from_slice` and `memcpy` now
  called out of line in the workbook-stream code (an inlining change; the same
  61 allocations). +7.0% user cycles in the counters.
- **`ppt_fresh_write_to` payload-heavy: 0.792 and 0.805** (a speed-up) in the
  harness configuration of 5 + 40 samples, but with 200 samples per process
  (`ppt-long/`) the base's per-process p50s are 0.625–0.659 ms and the branch's
  0.673–0.697 ms (+7%), and the pinned-threshold check shows the branch
  +3.5% in user instructions where the default run shows −2.7%. The probe
  (`write_to` alone) measures 0.997 with +0.1% instructions and identical
  allocations. The PPT change is structural (same records, same reservation),
  so these harness movements follow heap layout, which the harness's own
  per-process buffers (sized from `--samples`) change; none is claimed.
- `ppt_fresh_write_to` large: 1.023/1.024 wall, +0.6% user instructions, +14.4%
  kernel instructions (10,000 per iteration); the probe measures 1.000 and
  −0.9% instructions.
- `xls_fresh_write_to` payload-heavy: 1.013/1.012 wall, +0.1% instructions;
  probe 0.997.
- Counter noise on unchanged or faster paths: `doc_fresh_write_to` tiny −20.9%
  user cycles with −0.3% instructions; control `xls_semantic_one_edit_save`
  kernel cycles −63.6% (tiny, a few thousand cycles). The control's timed
  region runs 0.991–0.997 in both windows; the editor's only change is the
  `?` after `encode_ptg_tokens`, which its cases do not reach.

## Findings outside this change

Found while auditing the fresh writer's string encoders, by this record and by
its review; not fixed here (the coordinator queues them). "Reproduced" means a
litchi round trip showed it; the rest are from reading the code.

- **Internal hyperlinks** (`writer/biff/worksheet.rs:910,916`, review,
  reproduced): the length is `chars().count()` while the text is written as
  UTF-16, so a sheet named `R😀` with `set_hyperlink(…, "internal:'R😀'!A1")`
  makes litchi's reader drop the link and that worksheet's cells. The record
  length is also computed from a wrapped `truncate_usize_to_u16(wide.len())`
  before the "exceeds BIFF8 length limit" check.
- **AutoFilter strings** (`writer/biff/worksheet.rs:678`, review, reproduced):
  the length is the UTF-8 byte count capped at 255 while the whole string is
  written, as UTF-16 when non-ASCII; `café` gets length 5 with 4 units written,
  and the reader drops the filter (`autofilter.rs:303`).
- **Font names** (`writer/formatting.rs:300–317`, review): cut to 31 UTF-16
  units without a refusal.
- **Font index 4** (review, reproduced): a font added through `add_cell_style`
  gets logical index 4, which litchi's reader refuses ("Font logical index 4 is
  invalid", `font.rs:362`).
- **Worksheet-name characters** (review, reproduced): `add_worksheet` accepts
  `[]:*?/\`, NUL, U+0003 and leading or trailing apostrophes; `"a/b:c"`, NUL
  and U+0003 make the reader refuse the whole workbook. The fix could reuse
  `validate_sheet_name` (`cell_values/structural.rs:127`).
- **XFEXT, STYLEEXT and CRN** (review, not verified): record lengths may wrap
  at 65,536 bytes or more, and the XF index wraps at 65,536 styles.
- **Custom number-format count** (this record, reproduced): the writer
  registers any number of custom formats; litchi's reader refuses more than
  218 `Format` records, of which the writer always emits 8, so a workbook with
  211 custom formats is written and then refused on open (210 open). The index
  space itself ends at 392 (229 formats).
- **Formulas are tokenized only when written** (this record, reproduced):
  `write_formula(…, "SUM(")` succeeds and every `write_to` then fails
  ("Mismatched parentheses") until the cell is overwritten. Unlike a number
  format, the cell can be replaced, so the writer is not stuck.
- **PivotTable records** (`writer/biff/pivot/codec.rs`): `SXVIEW`, `SXVD`,
  `SXVI`, `SXDI` and the cache strings set `cch` from `chars().count()` while
  writing UTF-16 for non-ASCII names, so a supplementary-plane character leaves
  the count one short; counts and record lengths use wrapping
  `truncate_usize_to_u16`, and `validate_pivot_table_config` checks no name
  length (MS-XLS bounds these names to 255 characters).
- **The harness corpus comment** in `write_fresh_xls` ("`WritableWorksheet`
  stores cells in a hash map, so use one string cell per worksheet") no longer
  holds; a multi-string XLS selector would measure the path this record changes.
  The harness is unchanged here so the two legs run identical workloads.

## What is left

- **One sort per string worksheet instead of two.** The table sorts each
  worksheet's string cells, and the emission sorts all of its cells again
  (`writer/core/stream/codec.rs`). Sharing one sort — the table sorting every
  cell of a worksheet that has strings, keeping that list, and the emission
  reusing it — removes the 20.0M-instruction string sort of the 80,000-string
  probe, but every worksheet with strings then holds its 16-byte-per-cell list
  from staging until its records are written, instead of one worksheet's list
  at a time: about +0.96 MB (+7.5%) peak on the probe's four worksheets. An
  alternative without that memory — ordering distinct strings by their first
  cell, `sheet << 48 | row << 16 | column`, and renumbering — saves most of the
  cost only when strings repeat. Either is its own measured change.

## What is not claimed

- No claim is registered; `performance_claim: none`.
- Excel's behaviour with strings between 32,768 and 65,535 characters, or with
  the refused inputs, is not tested (no Excel on this host); the limits follow
  MS-XLS and the crate's reader.
- Timings are warm, in-memory, single-CPU, on a shared host; the multi-string
  and payload-heavy results move with glibc heap state, as the pinned runs show.
- The DOC and PPT follow-ups make no performance claim; their movements above
  are layout and heap-state effects of unchanged work.

## Verification

At `a400341bdf`, the last production commit, with
`CARGO_TARGET_DIR=…/targets/0757` (`gates.txt`; after the test-only `8a2693a83e`,
fmt, clippy and the litchi-xls tests were run again):

- `cargo fmt --all --check` — 0
- `cargo check -p litchi-xls -p litchi-doc -p litchi-ppt --all-targets --locked` — 0
- `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked` — 0
- `cargo clippy -p litchi-xls -p litchi-doc -p litchi-ppt --lib --no-deps --locked -- -D warnings` — 0
- the same with `--all-targets` — 0 for `litchi-doc` and `litchi-ppt`; 101 for
  `litchi-xls` on `tests/xls_query_index_cache.rs:546` (`unusual_byte_groupings`),
  which fails identically on the base (checked in the before worktree, as 0746
  and 0753 also recorded); with that lint allowed — 0
- `cargo test -p litchi-xls` — 1,536 passed at `8a2693a83e` (base 1,506); `-p litchi-doc` —
  1,206 (base 1,200); `-p litchi-ppt` — 1,235 (base 1,231); `-p litchi-ppt
  --features encryption` — 1,246 (base 1,242); `-p litchi --features
  doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` — 382
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-xls -p litchi-doc -p litchi-ppt --no-deps` — 0
- `python3 tools/check_crate_boundaries.py`, `python3 tools/non_iwork_gate.py verify`,
  `python3 tools/check_perf_claims.py … --mode structural` — 0
- `tests/xls_writer_text_goldens.rs` in twelve separate processes — 12 × 5
  passed; copied onto the base — the pre-0753 goldens pass, the multi-string
  pin and the order test fail ("many_strings is not deterministic")
- the release harness builds at `a400341bdf` (instruction check above)

After the review fixes (`6cebd664e6`), with a fresh
`CARGO_TARGET_DIR=…/targets/0757` (`gates.txt`, last section): fmt, `cargo
check` of the three crates, the facade and `tools/perf-baseline` (all
targets), clippy `--lib` and `--all-targets` (the same pre-existing
`unusual_byte_groupings` failure only) and rustdoc — 0; tests: `litchi-xls`
1,539, `litchi-doc` 1,206, `litchi-ppt` 1,235, `litchi-ppt --features
encryption` 1,246, the facade 382 — all passed; the boundary, non-iWork and
claims scripts — 0.

The harness did not change and calls none of the APIs made fallible, so its
tests and the coverage validator were not required.

## Cleanup

Recorded in `results/change-0757/cleanup.json`: the before-leg worktree
`0757-before-src` (removed with `git worktree remove --force`), the target
directories `targets/0757`, `0757-before`, `0757-final`, `0757-probe-A` and
`0757-probe-B`, and `scratch/0757/*` (binaries, probe projects, callgrind
outputs, logs, development measurements). Only summaries, raw JSON reports,
`perf stat` and callgrind logs, the probe sources and the scripts are kept in
the packet.
