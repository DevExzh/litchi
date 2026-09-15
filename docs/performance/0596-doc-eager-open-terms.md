# 0596: the eager DOC open stops decoding text four times, resolving each PAPX three times, and copying the WordDocument stream twice

Status: retained. `performance_claim: none` — this record carries deterministic
instruction and call counts on three real fixtures, native `perf stat` cycles
beside them, and paired harness timing with an A/A floor, all reported as
evidence rather than registered as a claim. A 15% regression the first
candidate caused on one scenario is reported in full, along with what removing
it showed about the harness.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements items **DOC-2, DOC-3 and DOC-4** (rank 11) of change
[0587](0587-remaining-opportunity-survey.md), measured against the driver and
method that record retained in `results/change-0587/doc-ppt/`.

## What was changed

Three independent terms of `Document::from_ole_with_options`
(`crates/litchi-doc/src/document/package.rs`), which is what
`litchi_doc::Package::document()` runs and what both validation passes of every
DOC edit-and-save pay.

### DOC-2: `TextExtractor::new` walked the text four times (`parts/text.rs`)

It decoded every piece into an unsized `Vec<u16>`, one `encode_utf16` plus
`extend_from_slice` per ANSI character; then `String::from_utf16_lossy`; then
`text.encode_utf16().count()` to size the CP map; then a `char_indices` walk to
fill it. Now:

- `extract_text_from_piece_descriptors` resolves each piece's clipped byte run
  once through a new `piece_run` helper, sums the exact code units those runs
  yield, and reserves that much. The buffer never grows, and the reservation is
  bounded by the bytes the `WordDocument` stream actually holds rather than by
  the character counts the piece table claims, so a hostile table cannot inflate
  it.
- ANSI runs decode through `WINDOWS_1252_TO_UTF16`, a 256-entry `u16` table
  built in a `const` block from the same `windows_1252_to_char` match. Every
  Windows-1252 byte maps to one Basic Multilingual Plane scalar, so one `u16`
  per byte is exact; a test asserts the table equals `encode_utf16` for all 256
  byte values. `extract_text_simple` uses the same table and finds its null
  terminator with `position` before reserving.
- `decode_with_cp_map` replaces `from_utf16_lossy` + `encode_utf16().count()` +
  `char_indices` with one `char::decode_utf16` pass that pushes the string and
  the CP map together. `decode_utf16` yields one scalar per valid surrogate pair
  and one `U+FFFD` per unpaired surrogate, so the decoded text spans exactly
  `utf16.len()` code units and the map holds exactly `utf16.len() + 1` entries —
  both sized once, neither counted twice. The old map builder is kept under
  `#[cfg(test)]` as the reference the new pass is differentially tested against,
  including unpaired surrogates on both sides and a pair split across a run
  seam.

### DOC-3: each PAPX entry was resolved about three times (`parts/pap_bin_table.rs`, `parts/pap/parser.rs`)

Per entry the old path parsed the same `grpprl` three times — once for the huge
PAPX / `PrcData` check, once in `apply_direct_sprms`, once more in
`from_sprm_with_stylesheet` to copy five table fields — allocated a `HashSet` per
entry, and re-resolved the style baseline whenever consecutive paragraphs
alternated styles. Now:

- `parse_properties_with_direct_cached` parses the entry once at the top. Both
  arms already opened by parsing `direct_sprms`, so hoisting it changes nothing
  about which error is raised first.
- `expand_data_indirections` splits into a recursive entry point that checks the
  depth limit and *then* parses, and `expand_parsed_indirections` that takes the
  parse. The split is what keeps a too-deep chain reported as too deep rather
  than as whatever its innermost `grpprl` happens to be.
- When no indirection expanded and the piece contributes no modifier, the
  concatenated `direct_grpprl` reproduces `direct_sprms` byte for byte, so that
  parse is handed to `cascade_styles_from_resolved_baseline` through a new
  `pre_parsed` parameter and reused by `apply_direct_sprms`.
- `apply_direct_sprms` derives the table state with `from_sprm_parsed` — the
  body of `from_sprm_context` over the parse it already holds — instead of
  parsing the same bytes a third time.
- `StyleBaselineCache` replaces change [0051](changes/0051-doc-adjacent-style-baseline-cache.md)'s
  single `(istd, baseline)` slot with a map of at most 64 entries and
  least-recently-used eviction. `resolve_style_baseline` is a pure function of
  the style index and the immutable stylesheet, so a baseline resolved for one
  run is the baseline every later run with that index would resolve; only
  successful resolutions are retained, exactly as 0051 required, and a direct
  `sprmPIstd` still never re-keys the cache. Because the most recently used
  entry is never the eviction victim, no input resolves a baseline more often
  than the one-entry cache did.
- One `PrcData` visit set is cleared per entry instead of constructed per entry.

### DOC-4: the FIB owned a copy of the `WordDocument` suffix (`parts/fib.rs`)

`FileInformationBlock::parse_at` did `data.to_vec()` over the whole stream
suffix, and the attached glossary FIB copied it again — on the 1.6 MB fixture,
697,827 bytes per open. The struct now holds `stream: Arc<Vec<u8>>` and
`offset: usize`, and a private `data()` yields `&stream[offset..]`.

The audit of every reader decided the shape. `raw_data()` is public, and its
bytes flow into `smart_tags::Snapshot::fib_bytes`, that snapshot's `finish()`
output and its `fingerprint()`; several validators compare an offset against
`fib.raw_data().len()`. **Truncating the FIB to its declared extent would
therefore move a refusal and change a public byte sequence, so this record does
not do it.** The shared view keeps `raw_data()` byte-identical instead: every
reader, including the fingerprint, sees exactly what it saw before. The
differential digest below checks that on every fixture.

`parse` and `parse_at` keep their `&[u8]` signatures and still copy, so the
~40 callers that hold borrowed bytes are unchanged. A new `pub(crate)
parse_shared` takes the shared stream; `Document::from_ole_with_options` and
`AttachedGlossary::parse` use it. `Document.word_document` becomes
`Arc<Vec<u8>>` and its public `word_document()` accessor still returns `&[u8]`.
The encrypted path decrypts into a private copy and re-shares it, so it pays one
stream copy where it used to pay two FIB copies. `Debug` is now written out by
hand so that it still prints the FIB's own bytes rather than the shared stream.

A fourth, smaller change came out of the measurement: `get_all_subdoc_ranges`
reads the eight `FibRgLw97` character counts once through `character_counts`
instead of resolving the FIB bytes once per count, because the shared view made
each resolution slightly more expensive. The seven public range accessors are
untouched; the prefix sums use `Sum for u32`, which is the same left-to-right
`+` chain they spell out.

## Why it is sound

Every step is value-identical by construction, not by adjustment.

- **Text.** The Windows-1252 table is a `const` transcription of the existing
  match, asserted equal to `encode_utf16` for all 256 inputs. The fused decode
  is `from_utf16_lossy` and the old map builder fused; a test compares it
  against both on seven inputs including lone surrogates. Piece clipping,
  truncation to whole UTF-16 pairs, the skip of descending and empty ranges and
  the skip of pieces starting past the stream are moved into `piece_run`
  unchanged. The lossy surrogate handling that `text_at_range` depends on is
  untouched, and `cp_to_byte` holds the same values.
- **Properties.** `parse_sprms` is a pure function of a byte slice, so handing a
  parse of the same bytes to a second consumer cannot change what that consumer
  sees. The `consumed != len` check still runs in both places and in the same
  order, so `"PAPX grpprl does not contain a whole number of SPRMs"` still
  precedes `"PAP grpprl does not contain a whole number of SPRMs"`. The reuse is
  taken only when the concatenated buffer is byte-identical to the parsed one.
  The depth limit is still checked before the parse it guards.
- **FIB.** `raw_data()` returns the same bytes at the same length; nothing reads
  past the FIB's extent that did not before, and nothing reads less.
- **Error identity.** No error site was added, removed, reordered or reworded.
  The differential digest records the exact `Debug` text of every refusal.
- **ADR reading.** ADR 0006: output bytes and preserved structures are
  unchanged, and the DOC writer is not touched. ADR 0003: no limit moved; the
  two new allocations are bounded — the UTF-16 reservation by the stream's own
  bytes, the baseline cache by 64 entries. ADR 0005: no new `unsafe`, no new
  dependency, no thread, no global, no ambient I/O, no public archive type,
  lock or executor. Retention falls: the eager `Document` held the
  `WordDocument` stream twice and now holds it once.
- **Untouched contracts.** The DOC writer and codec; `body_text` and its
  source-backed snapshot; `smart_tags` and every other part parser (they call
  `FileInformationBlock::parse` on borrowed bytes and still get an owning FIB);
  the encryption path's semantics; every public signature except the internal
  `cascade_styles_from_resolved_baseline`.

Record [0161](changes/0161-doc-public-validation-borrow-rejected.md)
rejected borrowing the owned editor's second validation; nothing here revisits
that.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0, CPU 16 pinned with `taskset`, `--release --locked`, other
agents active on the host. Base `08d968f8ec7db27cf1187d01911fd08b9d014d91`.
Binary sha256s are in
[`results/change-0596/README.md`](results/change-0596/README.md).

### Instructions per open (callgrind isolation pairs, s=10 against s=110)

Driver: the `docppt-survey` `doc-facade-open` mode retained by change 0587,
rebuilt against each leg.

| fixture | before | after | delta |
| --- | ---: | ---: | ---: |
| `saved-by-table.doc` (65,024 B) | 5,609,212 | 4,481,165 | **-20.11%** |
| `FloatingPictures.doc` (335,360 B) | 3,471,280 | 3,080,182 | **-11.27%** |
| kwsymphony form (1,619,457 B) | 4,971,153 | 4,058,410 | **-18.36%** |

The fourth fixture the brief named, `picture.doc`, refuses the eager open at
`Corrupted("invalid stylesheet: style names and aliases must be unique")` on
both legs and has no open to profile.

Where it went, inclusive per open. "outside the top 40" means the symbol is not
among the forty largest inclusive deltas of that profile, which for these opens
means under roughly 2% of the total:

| term | `saved-by-table` | `FloatingPictures` | kwsymphony |
| --- | --- | --- | --- |
| `TextExtractor::new` | 2,382,695 (42.48%) → 1,717,448 (38.33%) | 788,029 (22.70%) → 563,800 (18.30%) | 138,003 (2.78%) → outside the top 40 |
| `PapBinTable::parse` | 2,109,966 (37.62%) → 1,682,103 (37.54%) | 997,528 (28.74%) → 875,397 (28.42%) | 1,114,923 (22.43%) → 937,197 (23.09%) |
| `apply_direct_sprms` | 569,594 (10.15%) → 359,707 (8.03%) | 375,295 (10.81%) → 253,321 (8.22%) | 577,957 (11.63%) → 438,462 (10.80%) |
| `resolve_paragraph_style_sprms` | 183,471 (3.27%) → outside the top 40 | outside the top 40 | outside the top 40 |
| `FileInformationBlock::parse_at` | outside the top 40 | outside the top 40 | 698,280 (14.05%) → **outside the top 40** |
| `__memcpy_avx_unaligned_erms` | 1,067,341 (19.03%) → 982,273 (21.92%) | 487,597 (14.05%) → 432,896 (14.05%) | 1,059,610 (21.32%) → 349,836 (8.62%) |

Call counts per open:

| count | `saved-by-table` | `FloatingPictures` | kwsymphony |
| --- | --- | --- | --- |
| `parse_sprms` | 1,931 → **686** | 965 → **531** | 795 → **467** |
| PAPX entries (`concat`) | 626 → 626 | 475 → 475 | 391 → 391 |
| `parse_sprms` per entry | 3.08 → **1.10** | 2.03 → **1.12** | 2.03 → **1.19** |
| `resolve_paragraph_style_sprms` | 103 → **12** | 12 → 8 | 16 → **5** |
| `ParagraphProperties::clone` | 548 → 457 | 225 → 221 | 165 → 154 |
| `memcpy` | 10,722 → **8,955** | 7,349 → **6,615** | 8,754 → **7,438** |
| `malloc` | 4,524 → **3,069** | 3,122 → **2,678** | 2,907 → **2,526** |

DOC-3's falsifier was "`parse_sprms` calls per entry do not fall to one": they
fall to 1.10, 1.12 and 1.19. The remainder is the stylesheet's own sprms in
`paragraph_style_on_baseline`, nested `PrcData` expansions, and entries with an
empty `grpprl`. DOC-2's falsifier was "that fixture's open falls by under 10%":
`saved-by-table.doc` falls 20.11%.

### Cycles per open (native `perf stat`, median of three runs of 2,020 opens)

Callgrind counts `rep movsb` per byte, so the copy removals are priced natively
as well.

| fixture | instructions | cycles | task-clock |
| --- | ---: | ---: | ---: |
| `saved-by-table.doc` | 5,749,093 → 4,728,617 (**-17.75%**) | 1,633,987 → 1,415,946 (**-13.34%**) | 370 µs → 318 µs (-14.16%) |
| `FloatingPictures.doc` | 3,000,179 → 2,829,208 (-5.70%) | 854,248 → 871,077 (**+1.97%**) | 192 µs → 196 µs (+2.23%) |
| kwsymphony form | 5,674,117 → 4,552,121 (**-19.77%**) | 2,314,943 → 1,824,307 (**-21.19%**) | 519 µs → 409 µs (-21.25%) |

`FloatingPictures.doc` does 5.70% fewer instructions and 1.97% more cycles.
That is below the 5% review threshold, but it is instructions and cycles
disagreeing in sign on the same operation, which is the same phenomenon the
`paragraph_count` scenario shows below at a much larger size. This record does
not explain either.

### Paired harness timing

Harness: `tools/perf-baseline`, generated writer corpora (`tiny`, `large`), CPU
16, 30 warmups and 500 samples per leg, legs in the order A1 B1 B2 A2 with the
A/A floor taken as A3 against A4 in the same window, A and B pooled to 1,000
samples per state per window. Two windows were run against the committed
candidate (windows 3 and 4); two earlier windows against a superseded candidate
are discussed under *The regression* below. All figures are p50 unless stated.

| scenario | window 3 | window 4 | A/A floor (w3, w4) |
| --- | ---: | ---: | ---: |
| `doc_semantic_open` / tiny | **-4.70%** | **-5.43%** | -0.71%, +0.56% |
| `doc_semantic_open` / large | +1.47% | **-7.69%** | -2.42%, -2.54% |
| `doc_semantic_one_edit_save` / tiny | -2.39% | -0.98% | -1.39%, +1.33% |
| `doc_semantic_one_edit_save` / large | **-3.55%** | -2.23% | -0.82%, -1.48% |
| `doc_semantic_full_text` / tiny | +33.33% | 0.00% | 0.00%, 0.00% |
| `doc_semantic_full_text` / large | 0.00% | -1.47% | 0.00%, +4.23% |
| `doc_semantic_paragraph_count` / tiny | -4.55% | 0.00% | +4.76%, 0.00% |
| `doc_semantic_paragraph_count` / large | -2.63% | -0.32% | +0.04%, -0.25% |

Reading the table honestly:

- The `tiny` `full_text` scenario runs in 30 ns against a 10 ns timer tick, so
  its +33.33% is one tick. It is listed because the method says to report every
  scenario, not because it means anything.
- `doc_semantic_open/large` moves +1.47% and -7.69% across two windows whose
  floors are -2.4% and -2.5%. The two windows do not agree, so the only claim
  this record makes about the open's latency comes from the counts and from
  `perf stat` above, not from this row.
- The means and the p95/p99 columns of windows 2 and 4 are contaminated: seven
  other agents were building and measuring on the host, and those windows'
  own A/A floors reach -80% at p95 and -84% at p99. Windows 1 and 3 have clean
  tails. Full per-metric tables for all four windows are in the packet.
- The one direction that is consistent across all four windows and both shapes
  is `doc_semantic_one_edit_save`, which pays the open in both of its validation
  passes: -0.14% to -4.88% at p50, never positive.


### The regression that was found, and what fixing it demonstrates

The first candidate build — everything above except the
`get_all_subdoc_ranges` change — was **13% to 15% worse at p50 on
`doc_semantic_paragraph_count`, on both corpus shapes, in two independent
windows**, against an A/A floor of 0.3% or less at p50 for that scenario
(`timing/window1-summary.txt`: +15.45% large, +13.64% tiny;
`timing/window2-summary.txt`: +15.05%, +13.64%). That is far above the 5%
review trigger, and it was reproducible, so it was measured rather than
averaged away.

The counts said the scenario was not doing more work. On `saved-by-table.doc`,
with the document opened once and only the query timed:

- `Document::paragraph_count()` cost 1,650,265 → 1,650,767 instructions per
  call, **+0.03%**.
- Natively it cost 1,654,851 → 1,654,633 instructions and 257,914 → 256,199
  cycles per call, **-0.66% cycles**.
- 99.98% of the query is `ParagraphExtractor::count_paragraphs_in_range`, which
  walks `text.chars()` and is **not modified by this change**. Its instruction
  count moved by +464 out of 1.65 million; in the candidate build the compiler
  inlines it into the fold over the subdocument ranges, in the control build it
  stands alone.

The one thing the change did add to that path was per-count work in the FIB:
`get_all_subdoc_ranges` calls `get_character_count` about thirty times, and each
of those now resolved the shared stream instead of indexing an owned `Vec`.
`character_counts` resolves it once. Measured against the control, that brings
the whole query from +502 to +412 instructions per call and
`get_all_subdoc_ranges` itself to 299 → 289 instructions per call
(`counts/an-*-ranges.txt`) — **ninety instructions out of 1.65 million**. With
that edit in place the same scenario measures **-2.63% and -0.32% at p50** in
windows 3 and 4, and **0.00% and -4.55%** on the tiny corpus.

Ninety instructions moved this scenario by fifteen percentage points. That is
the finding worth carrying forward: on this host, at this scenario's size,
harness percentiles are dominated by where the recompiled loop lands, not by the
work it does. The record keeps all four windows so the swing is visible rather
than implied, and it does not claim the final -2.63% any more than it would have
claimed the earlier +15%.

## Correctness evidence

**Differential oracle over every `.doc` fixture in the tree.** A probe (retained
at `results/change-0596/probe/main.rs`) walks all 57 `.doc` files under
`test-data/` and prints, per file: the length and hash of `Document::text()`; the
number of paragraphs and a hash of the `Debug` form of every one of them, which
carries each paragraph's resolved `ParagraphProperties`, its runs and each run's
`CharacterProperties` and revision marks; `paragraph_count()` from the separate
public query; the section count and a hash of the section table; a hash of
`get_all_subdoc_ranges()`; and the length and hash of `FileInformationBlock::raw_data()`.
Files that refuse print the exact `Debug` text of the refusal instead.

**The two outputs are byte-identical** (`differential/digest-diff.txt` is empty).
42 fixtures open and 15 refuse, with all 15 refusals unchanged verbatim —
`PasswordRequired`, `UnsupportedVersion { nfib: 101, name: "Word 6.0" }`,
`UnsupportedVersion { nfib: 104, name: "Word 95 (7.0)" }`,
`Corrupted("invalid stylesheet: style names and aliases must be unique")`,
`Corrupted("grffldEnd.fNested disagrees with field containment")`,
`Corrupted("grffldEnd.fHasSep disagrees with the FieldList")` and
`Ole(InvalidFormat("Invalid byte order"))`. The FIB hash being identical is the
direct check that DOC-4 did not change a single byte any reader sees.

**Unit tests.** `cargo test -p litchi-doc`: 1,174 tests pass across 41 suites
(996 of them the crate's lib tests), 0 failed, 13 ignored.
The existing `parts/pap` tests are unchanged. `parts/text` gains
`windows_1252_table_matches_the_scalar_conversion` (all 256 bytes),
`fused_decode_matches_lossy_decode_and_the_reference_map` (seven inputs,
including lone high and low surrogates and a well-formed pair) and
`ansi_and_unicode_pieces_decode_as_before` (all 256 ANSI bytes, an odd-length
Unicode piece, a piece past the end). `parts/pap_bin_table`'s
`adjacent_style_cache_matches_scalar_cascade_and_rekeys` is extended to assert
the cache's resident key set and that a direct `sprmPIstd` still never re-keys
it, and `style_baseline_cache_is_bounded_and_evicts_least_recently_used`
asserts the bound and that the freshly used entry survives eviction.

**Writer output.** `doc_semantic_noop_edit_save` and `doc_semantic_one_edit_save`
verify exact no-op bytes, deterministic changed bytes, public reopen, forward
patch application and inverse restoration outside timing; the edit-save legs ran
1,000 samples per state per window with no verification failure.

**Gates.** `cargo fmt --all --check`, `cargo clippy -p litchi-doc --all-targets`,
`cargo test -p litchi-doc` and `cargo doc -p litchi-doc --no-deps` all pass;
tails in `results/change-0596/gates.txt`.

## Validation preserved

No validation was moved, weakened or skipped. The huge-PAPX / `PrcData` refusal
still fires from the same parse in the same place; the `PrcData` depth limit is
still checked before the `grpprl` it guards is parsed; the cycle detection still
uses a per-entry visit set (now cleared rather than rebuilt); the exact-SPRM-
consumption checks still run in both `apply_direct_sprms` and the table-state
derivation; `GlossaryMetadata::validate_text_boundaries` still runs against the
freshly built extractor; the FIB's magic check, minimum size, declared
`FibRgFcLcb` count and truncation refusals are untouched. Validation still does
not mutate.

## Limitations

- **Not claimed:** any speedup. `performance_claim: none`, no claim-registry
  entry. The counts are deterministic; the timings are two-binary comparisons on
  a shared host, and this record's own evidence shows a ninety-instruction edit
  moving one of them by fifteen points.
- **Not measured:** cold page cache, peak RSS, allocation profiles, any
  non-Linux or non-x86-64 target, encrypted documents (the encrypted path is
  exercised only by the two `PasswordRequired` fixtures, which refuse before the
  FIB is re-parsed), and documents with an attached glossary FIB — no fixture in
  the corpus has `pnNext != 0`, so DOC-4's second copy removal is reasoned from
  the source, not observed.
- **Unexplained:** the `FloatingPictures.doc` cycle increase (-5.70%
  instructions, +1.97% cycles natively), and the sign disagreement on
  `doc_semantic_open/large` between windows 3 and 4. Both are reported above
  with the counts that bound them. The `doc_semantic_paragraph_count`
  regression the first candidate showed is resolved in the committed code, but
  its cause — code layout, not work — is not something this change controls.
- **Left on the table:** the fused decode is still the largest single term of
  the `saved-by-table.doc` open — `TextExtractor::new` is 38.33% inclusive and
  1,336,150 instructions self, most of it filling `cp_to_byte`, which is still
  eight bytes per UTF-16 code unit and still built eagerly. Making it lazy, or
  narrowing its element type, changes `cp_to_byte_shared`'s crate-visible type
  and belongs in its own record. The per-entry `[sprms, piece_modifier].concat()` still
  allocates once per PAPX entry (626, 475 and 391 times per open) because the
  result is retained in the run; removing it means changing `ParagraphRun`'s
  representation. `TapParser` still re-parses the same `grpprl` for table
  paragraphs.
- **Not attempted:** truncating the FIB to its declared extent. The audit found
  public readers of `raw_data()` — `smart_tags::Snapshot::fib_bytes`, its
  `finish()` and its `fingerprint()` — so truncation is a contract change and
  needs a frozen design record of its own before anyone tries it.

## Retained evidence

[`results/change-0596/README.md`](results/change-0596/README.md) — callgrind
annotations for both legs of every fixture and every isolated read, the native
`perf stat` runs, all timing legs with their raw samples, the differential
digests, the probe source and every capture script.
