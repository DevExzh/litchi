# 0573: prove a ZIP local header in one read

Status: retained. `performance_claim: none` — this record carries deterministic
read counts only. **No timing is measured and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was removed

`validate_reader_entry_layout_with_name_policy` in `soapberry-zip` proved one
entry's strict local layout with **two** positional reads and one heap
allocation, for every entry, every time:

```rust
reader.read_exact_at(&mut fixed_buffer /* 30 bytes */, entry.local_header_offset)?;
// …parse, bounds-check…
variable_data.try_reserve_exact(variable_length)?;      // one Vec per member
reader.read_exact_at(&mut variable_data, local_header_offset + 30)?;
```

The strict-layout proof is whole-archive: `build_strict_layout_proof` runs it
over every member of the package before the first payload byte of the one member
the caller actually asked for. Change
[0567](0567-ooxml-single-index-per-open.md) identified that shape and change
[0572](0572-ooxml-range-source-attribution.md)'s plan priced it — predicting 264
requests for the 132-member `ConditionalFormattingSamples.xlsx`, and more than
85 percent of the cell-read phase on a latency-bearing source.

The two reads are now one speculative bounded read into a fixed-size stack
buffer, with the heap buffer and the second read kept as the fallback for a
variable region the window did not cover.

## How the window is sized and bounded

Two separate numbers decide the read length, and it is worth keeping them apart.

**The ceiling** is `STRICT_LOCAL_HEADER_WINDOW`, the stack buffer, **640 bytes**.
A local variable region can be up to `2 * u16::MAX` bytes, so no ceiling removes
the fallback; the question is only how often the fallback fires. The corpus
survey retained with this record answers it:

| window | OOXML members proved in one read | fall back |
| --- | ---: | ---: |
| 512 | 3,969 / 4,270 | **301** |
| 576 | 4,270 / 4,270 | 0 |
| **640** | **4,270 / 4,270** | **0** |

The largest local variable region in `test-data/ooxml` is 539 bytes, so the
smallest sufficient window is 569. The distribution is bimodal because Microsoft
Office writes a `0xa220` growth hint whose payload is 260 or 516 bytes, which is
why the obvious 512 is measurably too small: it would miss 301 of 4,270 members,
including two of the 132 in `ConditionalFormattingSamples.xlsx`. 640 covers every
member with 71 bytes of headroom. Across all 14,744 members of all 533 ZIP
containers in `test-data`, exactly two — the `xl/revisions/userNames.xml` member
of two revision-bearing XLSX fixtures — carry a 2,056-byte extra field and would
need a 2,112-byte window. They are the corpus's proof that the fallback is
reachable by real files and must stay.

**The bound** is where the member's own payload must begin. An entry that will be
accepted satisfies

```
local_header_offset + 30 + variable_length + compressed_size + descriptor
    <= next_local_header_offset
```

because that is exactly what the `data_end_offset` and `span_end` checks already
require, plus the caller's own non-overlap proof. Reserving `compressed_size` and
the widest descriptor therefore **cannot truncate a variable region the two-read
form would have accepted**, and in a gapless archive the window ends exactly at
the variable region. Proving a layout still reads only framing bytes.

That property is not decoration. Two existing tests encode it —
`preservation_zip64_promotion`'s "indexing must remain metadata-only after
archive location" and `litchi-opc`'s
`topology_changed_replacement_reads_source_payload_once` — and both fail against
a window bounded only by the central-directory offset. Threading the next
member's local offset through as a read hint is what keeps them passing; it is a
hint only, and a wrong value can change how many reads are issued but never which
checks run, in which order, or which error is produced.

`next_local_header_offset` is free at both production call sites:
`IndexedArchive::build_strict_layout_proof` walks `self.layout` in order, and
preservation's `validate_local_span` already computes it as `local_end`.

## Measured effect

Reads and bytes for one whole-archive strict layout proof over every
DOCX/PPTX/XLSX fixture under `test-data/ooxml`, counted through a positional
reader that records every `read_at`.

| | fixtures | members | total reads | total bytes |
| --- | ---: | ---: | ---: | ---: |
| before | 168 | 4,089 | **8,326** | 514,578 |
| after | 168 | 4,089 | **4,385** | 517,371 |
| | | | **−47.3%** | **+0.54%** |

| format | fixtures | members | descriptor members | reads before | reads after | bytes before | bytes after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| docx | 19 | 273 | 10 | 556 | **293** | 44,552 | 44,647 |
| pptx | 56 | 2,322 | 65 | 4,709 | **2,452** | 243,369 | 244,976 |
| xlsx | 93 | 1,494 | 73 | 3,061 | **1,640** | 226,657 | 227,748 |

The measurement matches a model with no free parameters: a descriptor-free member
costs one read, a descriptor-bearing member keeps the three it already cost, and
`3,941 × 1 + 148 × 3 = 4,385`.

**159 of 168 fixtures read byte-for-byte the same bytes as before while halving
their reads.** No fixture issues more reads than it did.

| fixture | members | reads | bytes |
| --- | ---: | ---: | ---: |
| `xlsx/ConditionalFormattingSamples.xlsx` | 132 | 264 → **132** | 10,740 → **10,740** |
| `pptx/smartart-tdf134221.pptx` | 54 | 108 → **54** | 4,606 → **4,606** |
| `pptx/smartart-autoTxRot.pptx` | 53 | 106 → **53** | 5,013 → **5,013** |
| `pptx/activex/activex_checkbox.pptx` | 51 | 102 → **51** | 4,929 → **4,929** |
| `pptx/backgrounds.pptx` | 48 | 96 → **48** | 4,727 → **4,727** |

Change 0572's plan predicted 264 requests for `ConditionalFormattingSamples.xlsx`.
**That prediction is confirmed exactly**, and the after figure is 132, with the
bytes unchanged.

### Two independent instruments agree

Change 0572 counted the same proof from the other side of the crate boundary,
through a counting `litchi_core::ReadAt` handed to `litchi-opc`, over a
`git archive` extraction of `163ac1bd6`. Its **before** figures and the ones
captured here agree with no adjustment:

| fixture | 0572 (OPC-layer `ReadAt`) | here (ZIP-layer `ReaderAt`) |
| --- | --- | --- |
| `xlsx/sheet-names.xlsx` | 26 requests / 2,478 B | 26 / 2,478 |
| `xlsx/universal-content.xlsx` | 39 requests / 861 B | 39 / 861 |
| `xlsx/ConditionalFormattingSamples.xlsx` | 264 requests / 10,740 B | 264 / 10,740 |

0572's prose renders the `sheet-names.xlsx` row as "13 reads and 2,478 bytes",
pairing this change's *after* read count with its *before* byte count; its own
table on the same page gives the before figure as 26, which is what this capture
reproduces.

Change [0575](0575-zip-lazy-strict-layout-design.md) is a design record written
against this change and predicts its result for three fixtures independently, as
its row "d — change 0573". All three are confirmed by the capture here:

| fixture | 0575 predicts | measured |
| --- | --- | --- |
| `xlsx/ConditionalFormattingSamples.xlsx` | 132 reads / 10,740 B | 132 / 10,740 |
| `xlsx/sheet-names.xlsx` | 13 reads / 2,478 B | 13 / 2,478 |
| `pptx/shapes.pptx` | 48 reads / 4,748 B | 48 / 4,748 |

0575's model assumes one read per header throughout. Those three fixtures carry
no data descriptors, so the assumption holds for them; it does not hold for the
eight all-descriptor fixtures named below, and 0575's own limitations already
record that its corpus excludes descriptor-bearing archives.

### The byte increase, in full

Bytes rise by 2,793 across the corpus, 0.54 percent, and all of it lands on the
eight fixtures whose members *all* carry a data descriptor. Those members gain
nothing and pay their variable region twice:

| fixture | members | reads | bytes |
| --- | ---: | ---: | ---: |
| `pptx/shape-soft-edges.pptx` | 65 | 195 → 195 | 5,117 → 6,724 (+31.4%) |
| `xlsx/universal-content.xlsx` | 13 | 39 → 39 | 861 → 1,020 (+18.5%) |
| `xlsx/autofilter.xlsx` | 11 | 33 → 33 | 726 → 858 (+18.2%) |
| `xlsx/formats.xlsx` | 12 | 36 → 36 | 785 → 922 (+17.5%) |
| `xlsx/duplicate-defined-names.xlsx` | 11 | 33 → 33 | 710 → 826 (+16.3%) |
| `xlsx/named-ranges-global.xlsx` | 9 | 27 → 27 | 579 → 672 (+16.1%) |
| `xlsx/column_style.xlsx` | 8 | 24 → 24 | 509 → 586 (+15.1%) |
| `docx/comment.docx` | 10 | 30 → 30 | 635 → 730 (+15.0%) |

One more fixture, `xlsx/shared-strings-malformed-count.xlsx`, is mixed: 9 of its
14 members carry descriptors, so its reads fall 37 → 32 while its bytes rise
1,214 → 1,591. Every other fixture is byte-identical.

The mechanism is stated plainly under Limitations. In absolute terms the worst
fixture reads 1,607 more bytes. The largest ratio of proof bytes to file size is
28.93 percent, and it is **unchanged**: the fixture that sets it,
`xlsx/DataValidationListTooLong.xlsx`, is one of the 159 that are byte-identical.

### Allocations

The per-entry `Vec` for local metadata is gone from the fast path. Over this
corpus that is **4,089 allocations before, 148 after** — one per
descriptor-bearing member, none at all for the other 3,941. An allocation that
does not happen also cannot fail, so the `Allocation { resource: "strict ZIP
local metadata" }` error is now unreachable for a member proved in one read;
its position in the error order is otherwise untouched.

## Validation preserved

Every check keeps its position and its identity. The decisions worth recording:

- **The error order is unchanged**: metadata → `local_header_offset >
  central_directory_offset` → fixed-read I/O → signature → `local_end >
  central_directory_offset` → variable-read I/O → method → flags → name → size
  framing → sizes → CRC → span. The speculative read is deliberately allowed to
  come back short: only the 30 fixed bytes are required at that point, and a
  variable region the window did not reach is resolved *after* the `local_end`
  bounds check, so a variable-region I/O error can still never precede it.
- **A short source keeps `read_exact_at`'s exact error.** The window uses
  `try_read_at_least_at`, which returns the byte count instead of failing, and
  fewer than 30 bytes is turned into the same
  `io::ErrorKind::UnexpectedEof` / `"failed to fill whole buffer"` the two-read
  form produced. `read_at_least_at` was deliberately *not* used: it reports
  `ErrorKind::Eof`, which is a different error.
- **The window has a 30-byte floor.** A local offset inside the central directory
  passes the offset check but leaves less than a fixed header before it. The
  two-read form still read all 30 bytes and reported the signature; clamping the
  window below 30 would have reported an early end-of-file instead. The floor is
  the one place the window is allowed past the central directory, and a test
  pins it.
- **The fallback is byte-for-byte the historical path**: the same
  `try_reserve_exact` with the same `resource` label, the same `resize`, the same
  `read_exact_at(variable_length)` at `local_header_offset + 30`, and the same
  `variable_length != 0` guard.
- **A short-reading source falls back rather than guessing.** Any source that
  returns fewer bytes than the window asked for, for any reason, takes the
  fallback and therefore behaves exactly as it did before.
- **Preservation is unchanged.** `validate_preservation_entry_layout` shares the
  function with `allow_name_mismatch: true` and still reports the mismatch; a
  test proves both the mismatch report and the strict refusal.
- **The proof's own ordering guarantees are untouched.** The next-member offset is
  read from the layout the proof is about to validate, but the non-overlap and
  duplicate-span checks that actually validate that order are unchanged and still
  run after each entry.
- No `unsafe`; `deny(unsafe_code)` intact. Allocation stays fallible, and there is
  one less of it.

## Correctness evidence

Fifteen tests were added, of which **seven fail against the pre-change code**:

| fails pre-change | what it pins |
| --- | --- |
| `strict_layout_proves_an_ordinary_member_in_one_positional_read` | the request sequence is one read, not `(0, 30)` then `(30, 288)` |
| `strict_layout_reads_exactly_the_window_boundary_in_one_call` | a 610-byte variable region is one read; 611 bytes is two |
| `the_speculative_window_never_reads_this_members_payload` | the window ends exactly where the payload begins |
| `the_window_may_reach_past_this_member_but_stops_at_its_size` | a caller that passes the directory offset is still capped at 640 |
| `a_zip64_local_sentinel_resolves_from_the_single_read_window` | the `u32::MAX` sentinel path resolves out of the same read |
| `a_descriptor_bearing_member_still_reaches_the_descriptor_reader` | window plus descriptor read, where three reads were needed |
| `preservation_keeps_its_name_mismatch_policy_on_the_single_read_path` | preservation shares the single-read path |

The eight that pass on both sides are the point of the exercise rather than
filler — they are the error-identity contract, and each one would have caught a
plausible way of getting this wrong:

| passes on both sides | what it pins |
| --- | --- |
| `a_source_ending_before_the_fixed_header_reports_the_historical_eof` | four truncation points before 30 bytes, same `UnexpectedEof` and message |
| `a_source_ending_inside_the_variable_region_reports_the_historical_eof` | three truncation points inside the variable region, same error, resolved by the same second read |
| `the_directory_bounds_check_still_wins_over_a_truncated_variable_region` | both faults present at once: `Eof` still wins and no second read is issued |
| `a_local_offset_inside_the_central_directory_keeps_its_signature_error` | the 30-byte floor; a plain `min` clamp would report end-of-file here |
| `an_oversized_variable_region_still_takes_the_historical_second_read` | the 2,056-byte-extra shape from the real corpus still works |
| `a_final_narrow_descriptor_keeps_the_historical_two_reads` | the descriptor reservation degrades to the historical three reads |
| `a_short_reading_source_still_proves_the_same_layout` | six chunk sizes from 1 to 512 bytes all produce the identical layout |
| `strict_layout_checks_keep_their_order_around_the_single_read` | method before flags before name before sizes before CRC |

Gates run for this change: `rustfmt --check` clean on all three changed files —
the workspace-wide `cargo fmt --check` also reports a diff in
`crates/litchi-xls/src/records.rs`, which is another change's in-flight file and
not touched here; `cargo clippy --workspace --all-features --lib --no-deps -- -D
warnings` clean, which confirms change 0571's `litchi-opc` and facade repair
still holds; `cargo clippy -p soapberry-zip --all-targets` clean, which covers
the new tests as well (the crate does not opt into the workspace lint table, so
that run is explicit); `soapberry-zip` **570 tests, zero failures**;
`litchi-opc` **664 tests, zero failures**.

## Limitations

No timing, allocation-profile, cold-cache, physical-device or cross-platform
result is claimed. The counts are logical `ReaderAt` calls over warm fixtures. A
47 percent reduction in positional reads is not separable from timing noise on a
warm local file; change 0572 is the record that prices these requests against a
latency-bearing source, and nothing here asserts a wall-clock result.

Three gaps are recorded rather than worked around.

**Descriptor-bearing members gain nothing.** The window must reserve room for the
data descriptor, and the descriptor's encoded width — 12, 16, 20 or 24 bytes —
is not knowable until the local header has been parsed, which is the read being
planned. Reserving the widest form is the only bound that is sound for every
input, including a forced-ZIP64 local header on a ZIP32 central record, so a
member carrying a narrower descriptor falls back and pays its variable region
twice. Every descriptor in this corpus is narrower, which is why 148 members keep
three reads and why bytes rise at all. Narrowing the reservation using
`entry.zip64_sizes` would win those members back for the common signed ZIP32
form, but it would read up to eight payload bytes on the forced-ZIP64 local
shape, so it was not done.

**Gaps between members are read.** The bound is the next member's local offset,
not the current member's true end, so padding or alignment filler between members
is read up to the 640-byte ceiling. No corpus fixture has such a gap, so this is
modelled rather than measured.

**A caller that does not track member order gets a looser window.** Passing the
central-directory offset is always correct and always safe, but the window is
then capped only by 640 and can read later members' bytes. Both production
callers pass the real next-member offset; the looser case exists so the function
stays callable for a single entry, and a test pins its behaviour.

Finally, this change touches the reader-backed prover only. `ZipSliceArchive`
proves the same layout against an in-memory slice with no reads at all and is
untouched, so `office::ArchiveReader`, which is slice-backed, is unaffected by
construction.
