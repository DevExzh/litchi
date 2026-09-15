# 0583: a neighbour's ZIP span bound reserves the larger of its two declared payload lengths

Status: retained. `performance_claim: none` — this record carries no timing,
allocation, read-count or byte-count improvement, and claims none. It is a
correctness and safety fix to change
[0580](0580-zip-target-scoped-strict-layout.md), closing the defect change
[0582](0582-zip-strict-scope-differential-fuzz.md) recorded as **finding 2**.
The only quantitative result about cost is a *non-change*: change 0580's corpus
census is byte-identical.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was fixed

Change 0580 replaced an archive-wide ZIP strict-layout proof with one scoped to
the member being read. A record the caller is **not** reading is no longer
validated; it is only *bounded* — the proof asks where that record claims its
span ends, and refuses the read if the claim covers the target. For a
predecessor inside the residual window, the bound costs one 30-byte read of that
record's fixed local header.

`local_span_bound_from_fixed` in `crates/soapberry-zip/src/archive.rs` computed
that bound from three terms. Two came from the local header it had just read —
`file_name_length` and `extra_field_length`, so an inflated local name or extra
field was correctly caught. The third, the payload length, came from the
**central** record:

```rust
let payload_end = entry.local_header_offset
    .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
    .and_then(|offset| offset.checked_add(variable_length))
    .and_then(|offset| offset.checked_add(entry.compressed_size))   // central, always
```

`ZipLocalFileHeaderFixed::parse` has already decoded the local header's own
`compressed_size` at the top of that function, and the descriptor branch two
lines below reads it for the ZIP64 sentinel. It was never compared. A record
whose local header declares a longer payload than its central record therefore
declared a span the proof could not see, and a member lying inside that span was
readable.

The fix splits the two branches and takes the larger of the two declared lengths
on the one that is a pure threshold:

```rust
fn neighbour_payload_length(
    file_header: &ZipLocalFileHeaderFixed,
    entry: &ZipArchiveEntryWayfinder,
) -> u64 {
    if file_header.compressed_size == u32::MAX {
        return entry.compressed_size;              // the ZIP64 sentinel is not a length
    }
    entry.compressed_size.max(u64::from(file_header.compressed_size))
}

// in local_span_bound_from_fixed, after the variable region:
if entry.has_data_descriptor {
    let payload_end = variable_end.checked_add(entry.compressed_size)...;
    return Ok(LocalSpanBound::Descriptor { payload_end, local_zip64_sentinel });
}
let payload_end = variable_end.checked_add(neighbour_payload_length(&file_header, entry))...;
Ok(LocalSpanBound::Exact(payload_end))
```

Both records declare a length for the same physical region. The exact bound
reserves the larger. It costs **no read**: the field is decoded from 30 bytes the
caller already has in hand. Nothing else under `crates/` changed; `office.rs`
gained only test-module additions.

### Where the maximum does not apply, and why

**A declared data descriptor.** When the central record sets general-purpose bit
3, the bound is a `LocalSpanBound::Descriptor`, and its `payload_end` is not only
a refusal threshold — it is the offset `resolve_span_end` **reads the descriptor
at**. Moving it does not make the bound larger; it makes it *different*. The
resolved span end is `payload_end + encoded_size`, and `encoded_size` is
data-dependent (12, 16, 20 or 24 bytes), so it is not monotone in `payload_end`:
a shifted read can find a descriptor that matches where the true one did not.
That branch therefore keeps the central length, exactly as change 0580 computed
it.

This is not a theoretical concern. **It was a real class-C defect in an
intermediate implementation of this change, found in review and fixed before
anything was recorded.** An implementation that applied the maximum to the
descriptor branch turned a refusal into an acceptance on this archive:

```
pred.bin   @0   CENTRAL flags 0x08 (descriptor), central compressed_size 16
                LOCAL   flags 0x0000, LOCAL compressed_size 20
                payload 38..54, four filler bytes 54..58,
                a valid 12-byte unsigned data descriptor 58..70
target.bin @72  an ordinary Store member, 16 bytes
```

Resolving at 54, where the central length puts the payload end, finds filler and
fails, so the neighbour cannot be bounded and the read is refused. Resolving at
58, where the inflated local length points, finds a valid descriptor ending at 70
— two bytes clear of the target — and the read succeeds.
`a_descriptor_bearing_predecessors_payload_end_is_not_moved` is that archive, and
the `max-everywhere` mutant in `mutation-matrix.txt` is that implementation.

**The ZIP64 sentinel.** `u32::MAX` in a local `compressed_size` means "the real
value is in a ZIP64 extra field", which lives in the variable region this probe
deliberately does not read. Taking the maximum literally would reserve 4 GiB for
every ZIP64 member and refuse valid archives.
`a_local_zip64_size_sentinel_is_not_a_neighbour_span_length` pins it, and the
`sentinel-literal` mutant kills it.

Both fall back to the central length — exactly the bound change 0580 computed —
so neither weakens an existing bound.

### Order independence

The two fallbacks are also what keeps a verdict independent of read order, and
this too was measured rather than assumed. A record that validates as a *target*
has its exact span end memoised as its neighbour bound by
`StrictLayoutMemo::merge`, so the memoised value and the probed value must agree
or the same archive accepts or refuses a later member depending on which member
was read first. The invariant that makes them agree is:

> Every record that can be validated as a target has
> `neighbour_payload_length == entry.compressed_size`.

`validate_reader_entry_layout_with_name_policy` refuses a target whose local and
central **flags** differ before it compares anything else, so a target's
descriptor bit agrees in both records and a descriptor-bearing one takes the
excluded branch. For a non-descriptor target the size guard is unconditional, and
`resolve_local_entry_size_framing` substitutes the ZIP64 extra field's value only
for the field that actually carries `u32::MAX` — so a validated non-descriptor
target has raw local `compressed_size == entry.compressed_size`, or the sentinel,
which is the other fallback.

**The first implementation of this change did not have that invariant and the
harness said so**: applying the maximum with no branch split produced **244
intra-build order-independence oracle failures** on descriptor archives whose
local `compressed_size` had been mutated, across
`I.read_entry_to.order_independent` and `R.read_to.order_independent` on both
limit profiles. The shipped form produces **0**.
`a_descriptor_bearing_predecessors_local_size_is_not_a_span_length` reads the
predecessor first and then the target on one reader, and the target alone on a
fresh one, and requires the same verdict.

### Which direction is safe

The probed bound is consumed in exactly two places, both refusal predicates:

```rust
if bound.min_span_end() > target_offset { return Err(strict_overlap_error()); }
if bound.max_span_end() > target_offset {
    let span_end = resolve_at(position, bound)?;      // Descriptor branch only
    if span_end > target_offset { return Err(strict_overlap_error()); }
}
```

For an `Exact` bound, `min_span_end == max_span_end == payload_end` and
`resolve_at` returns it unchanged, so a larger `payload_end` can only turn a
non-refusal into a refusal. That is why the maximum is confined to that branch:
the `Descriptor` branch reaches `resolve_at`, which is a *read at an offset
derived from the bound* rather than a comparison against it, and is therefore not
monotone. The three early exits that prune a predecessor before any probe — the
`break` on the prefix maximum, the `continue` on `zero_io_max_span_end`, and the
refusal on `zero_io_min_span_end` — are computed from central metadata alone and
are untouched, so no pruning decision moves either.

The property was then measured, not just argued. `delta.py` pairs all three
builds on `(input, profile, API, member)`:

| transition, pre-change → 0580 → 0580+0583 | verdicts |
| --- | ---: |
| `err → ok → err` — change 0580 made it readable, change 0583 refuses it again | 32,472 |
| **anything → refuse → accept** — a read change 0580 refused that change 0583 admits | **0** |

Zero, over 1,924,524 member verdicts. Every one of the 32,472 verdicts the fix
moved moved toward refusal.

### Error identity

A member newly refused by this fix carries
`ErrorKind::InvalidInput { msg: "strict streaming refuses overlapping ZIP local
spans" }` — `strict_overlap_error()`, the identity change 0580 already uses when
a neighbour's declared span covers the target. That is the correct identity: the
refusal *is* an overlap, discovered in the one place change 0580 was not looking.
It is also an identity a caller already has to handle on this path, so nothing
new appears in the contract.

**No error identity is added, removed or renamed.** One uniform extraction over
all three full reports, with the name/method/reason payloads collapsed so the
comparison is about identities rather than messages, gives 20 distinct
`InvalidInput` identities on the pre-change build and 21 on both the 0580 and the
0580 + 0583 builds. The `before → 0580` diff is a single addition — change 0580's
now-reachable duplicate-offset refusal, which change 0582 already recorded. The
`0580 → 0580 + 0583` diff is **empty**. `error-vocabulary.txt` retains all three
lists and both diffs. (The absolute counts differ from change 0582's 39 and 40
because that record's extraction kept the embedded names; the delta is the same
and it is what this paragraph claims.)

### The uncompressed-size field: deliberately not treated

Change 0582 reports that `neighbour-local-usize-smuggles` also flips a verdict,
and classifies it as an ordinary finding-1 relaxation rather than a
payload-placement differential. **That classification is correct and this change
deliberately leaves it alone.**

`uncompressed_size` never describes an on-disk region. The bytes that separate
one local record from the next are `compressed_size` bytes long, whatever the
member inflates to (APPNOTE 4.3.7 and 4.4.8), so no reader — streaming or
central-directory-driven — can be made to place a payload boundary differently by
inflating it. There is no second parse of the same bytes to diverge from.
`a_predecessor_local_uncompressed_size_does_not_move_the_payload` pins the
decision and the `usize-max` mutant kills it, so the choice is tested rather than
merely asserted. The corpus agrees: all eight `neighbour-local-usize-smuggles`
target rows in `crafted-verdicts.txt` — four entry points, two limit profiles —
are byte-identical between change 0580 and change 0580 + 0583.

It remains an instance of change 0580's finding-1 relaxation — a local-versus-
central disagreement on a record the caller is not reading is no longer checked —
and that relaxation is change 0580's to own, not this change's.

## The witness, before and after

`crafted/neighbour-local-csize-smuggles.zip`, 388 bytes, two Store members:

```
pred.bin     local offset 0    central compressed_size 16   LOCAL compressed_size 100000
target.bin   local offset 200  central compressed_size 16   LOCAL compressed_size 16
```

A reader that follows local headers places `pred.bin`'s payload at
`[38, 100038)`, covering `target.bin`'s entire local record. A reader that trusts
the central directory places it at `[38, 54)`, which does not.

| target | pre-change (`93a610ded`) | change 0580 | change 0580 + 0583 |
| --- | --- | --- | --- |
| `pred.bin` | REFUSE `InvalidSize { expected: 16, actual: 100000 }` | REFUSE, same | REFUSE, same |
| `target.bin` | REFUSE `InvalidSize { expected: 16, actual: 100000 }` | **accept, 16 bytes** | **REFUSE, overlap** |

`crafted/neighbour-local-csize-smuggles-deflate.zip` is the same shape with
Deflate members: **480 decoded bytes of `target.txt`** that change 0580 returned
are refused again.

Both, on **all four** entry points that reach `strict_layout_for`
(`ArchiveReader::read_to`, `IndexedArchive::read_entry_to`,
`with_verified_entry_reader`,
`read_entry_precompressed_and_decoded_with_progress`), both readers, both limit
profiles — 16 verdicts. `crafted-verdicts.txt` carries all of them.

The neighbouring shapes are unmoved, which is the control that the maximum does
the one thing it was written to do:

| witness | change 0580 | change 0580 + 0583 |
| --- | --- | --- |
| `neighbour-local-csize-understates` (local 16, central 4096) | REFUSE, overlap | REFUSE, overlap |
| `neighbour-local-usize-smuggles` | accept | accept, unchanged |
| `neighbour-local-extra-inflated`, `predecessor-inflated-local-name`, `window-max-residual-*`, `predecessor-csize-reaches` | as change 0582 recorded | unchanged |
| every `descriptor-*` crafted archive | as change 0582 recorded | unchanged |

## Reachability, verified independently

Change 0582 claims the narrowed paths are reachable from `litchi-opc` and
therefore from the OOXML facades. That claim was re-verified against the working
tree rather than taken on trust. It holds, with corrections worth recording.

* **All five `litchi-opc` call sites reach `strict_layout_for`**, which lives in
  `crates/soapberry-zip/src/office.rs` (two copies, `:2825` for `ArchiveReader`
  and `:3913` for `IndexedArchive`) — not in `archive.rs`. It is **not gated**:
  `soapberry-zip` declares no cargo features, and `ArchiveValidationPolicy` never
  reaches the prover. One listed name is off by a hop —
  `source_backed.rs:3844` calls `read_entry_to_with_accounting`, which
  `read_entry_to` wraps.
* **A concrete OOXML chain exists.** `SourceBackedWorkbook::open` → `::sheet` →
  `SourceWorksheet::cell` → `::stream_cell`
  (`crates/litchi-xlsx/src/workbook/source.rs:683`) →
  `PartView::with_verified_decoded_reader`
  (`crates/litchi-opc/src/source_backed.rs:3610` → `:3682` → `:3773`) →
  `IndexedArchive::with_verified_entry_reader` → `strict_layout_for`
  (`office.rs:4068`). DOCX reaches it through
  `crates/litchi-docx/src/source_backed/tail_append_stream.rs:2036` and PPTX
  through `crates/litchi-pptx/src/presentation/source_cross_copy.rs:609`.
* **But the most common part read does not reach it.** `PartView::data()` goes
  through `read_part_with_observer_and_capture` to `load_part_with_accounting`'s
  non-capture branch (`crates/litchi-opc/src/source_backed.rs:10110`), which
  calls `read_entry_with_accounting` — a path that never enters the strict-layout
  proof. Change 0582's "reachable from the DOCX, XLSX and PPTX facades" is true;
  "reachable from every OOXML part read" would not be.
* `read_stored_borrowed` is structurally unreachable from `source_backed.rs` (it
  exists only on the slice-backed readers), and `preservation_index_with_limits`
  is reached exactly once, from the publication path at `source_backed.rs:9587`.
  Change 0582's characterisation of both as untouched archive-wide surfaces
  holds.

## Validation preserved

Change 0580's corpus-convergence result is reproduced exactly. Reading **every**
member of every DOCX/PPTX/XLSX fixture under `test-data/ooxml`, through change
0580's own census probe on the fixed build:

```
CENSUSTOTAL fixtures=168 members=4089 reads=4385 bytes=517371 accepted=4089
```

`diff` against change 0580's retained `census-after.txt` is **empty**, and so is
`diff` against its `scope-after.txt` — every per-fixture accept/refuse string,
read count and byte count is byte-identical, and so is the five-scenario read
table the change-0580 record publishes. The fix is a no-op on every archive in
`test-data`, in verdicts *and* in cost.

The differential says the same from the other side. Over change 0582's corpus:

| family | strict member verdicts, pre-change | change 0580 + 0583 |
| --- | --- | --- |
| `testdata`, `fuzz` | 54,023 accept, 0 refuse | identical |
| `testdata`, `wide` | 55,179 accept, 0 refuse | identical |
| `seeds`, both | 74 accept, 0 refuse | identical |

`per_family` in `classification-summary.json` has **no `testdata` or `seeds` row
in any divergence class**. Not one verdict changes on any of the 516 distinct
real archives or the 7 retained seeds.

The mechanism is the one change 0582 predicted: the pre-change archive-wide proof
accepted every member of every real archive, which required every record's local
and central sizes to agree, so the maximum is the central value everywhere in
`test-data`.

## Correctness evidence

Six tests were added to `crates/soapberry-zip/src/office.rs`, and the test-only
`ScopedMember` fixture builder gained `local_compressed_size`,
`local_uncompressed_size`, `local_flags`, `local_crc` and `descriptor_gap` so a
local header can declare sizes, flags and a descriptor position its central
record does not.

Run against the **pre-fix** tree — change 0580's `archive.rs`
(sha256 `df9ed280…`) with this change's `office.rs`, so the only difference is
the production fix — one fails. Five regression guards pass pre-fix by
construction, so each was also run against a mutant of the shipped fix. Every
test is killed by exactly one mutant and every mutant is killed:

| test | pre-fix (`prefix` mutant) | also killed by | what it pins |
| --- | --- | --- | --- |
| `a_predecessor_whose_local_compressed_size_reaches_the_target_is_refused` | **FAILS** — `unwrap_err()` on `Ok([116; 16])`, the 16 bytes the pre-fix build hands back | — | the defect |
| `a_descriptor_bearing_predecessors_payload_end_is_not_moved` | passes | `max-everywhere` | the resolve offset must not move |
| `a_descriptor_bearing_predecessors_local_size_is_not_a_span_length` | passes | `max-everywhere` | order independence through the memo |
| `a_predecessor_whose_local_compressed_size_understates_keeps_the_central_bound` | passes | `local-only` | the maximum runs in both directions |
| `a_local_zip64_size_sentinel_is_not_a_neighbour_span_length` | passes | `sentinel-literal` | the ZIP64 sentinel is not a length |
| `a_predecessor_local_uncompressed_size_does_not_move_the_payload` | passes | `usize-max` | the deliberate non-treatment of the uncompressed size |

`prechange-tests.txt` retains the pre-fix run and `mutation-matrix.txt` the five
mutants; `mutate.py` applies them.

Full suites, on a `git archive` extraction of `93a610ded` with the two files
overlaid and its own `CARGO_TARGET_DIR` — the six crates on the reachable path:

```
soapberry-zip:       passed=587  failed=0 ignored=2
litchi-opc:          passed=664  failed=0 ignored=1
litchi-ooxml-common: passed=255  failed=0 ignored=0
litchi-xlsx:         passed=1294 failed=0 ignored=0
litchi-docx:         passed=1449 failed=0 ignored=31
litchi-pptx:         passed=854  failed=0 ignored=2
```

`cargo fmt -p soapberry-zip -- --check` clean;
`cargo clippy -p soapberry-zip --all-targets -- -D warnings` clean. No `unsafe`,
no new allocation, no new read, no weakened bound: every existing refusal that ran
before still runs, and the `checked_add` chain is unchanged except for the term it
adds.

## The re-run

Change 0582's harness, corpus, classifier and limit profiles, unchanged, against
the fixed build. The corpus was verified byte-identical to change 0582's
fingerprint before reuse — 22,875 files, 334,625,602 bytes, the crafted family's
SHA-256s matching — so `RNG_SEED = 0x05820580` was not re-run. The pre-change
build and its report were reused for the same reason: `93a610ded` has not moved.

| outcome | change 0580 alone | change 0580 + 0583 | delta |
| --- | ---: | ---: | ---: |
| **panics, either side** | **0** | **0** | 0 |
| **intra-build oracle failures** | **0** | **0** | 0 |
| accept → refuse (class C) | 0 | 0 | 0 |
| accept → accept, different bytes (class D) | 0 | 0 | 0 |
| refuse → accept, pre-change refusal was overlap (class A) | 37,176 | 37,160 | **−16** |
| refuse → accept, pre-change refusal was something else (class B) | **1,094,221** | **1,061,765** | **−32,456** |
| refuse → refuse, different typed identity (class E) | 351,525 | 383,981 | +32,456 |
| identical | 1,209,509 | 1,209,525 | +16 |

**The 1,094,221 figure moves to 1,061,765, a reduction of 32,456**, and class A
loses 16 more, for 32,472 member verdicts that change 0580 made readable and
change 0583 refuses again. The shipped delta is 1,131,397 → 1,098,925 member
verdicts.

The two halves land differently, and the split is exact. The 32,456 that leave
class B become class E: the pre-change build refused them with some other
per-record identity, and they are now refused as the overlap they are. The 16
that leave class A become **identical** to the pre-change build — it refused them
as an overlap, and so does this one, with the same identity. That is why the
identical row moves by 16 and not by 32,472.

What the re-refused verdicts were refused for *before* change 0580 — the
`delta-summary.json` breakdown of all 32,472:

| pre-change refusal | verdicts |
| --- | ---: |
| `InvalidSize { expected, actual }` — local and central sizes disagree | 32,024 |
| `strict local and central flags differ` | 102 |
| `strict local and central names differ` | 102 |
| `strict local and central compression methods differ` | 96 |
| `stored local and central flags differ` | 34 |
| `stored local and central names differ` | 34 |
| `stored local and central compression methods differ` | 32 |
| `local ZIP64 size framing requires version 45 or newer` | 24 |
| `strict streaming refuses overlapping ZIP local spans` | 16 |
| `borrowed access refuses nonzero disk-start metadata` | 8 |

98.6% of the movement is the identity finding 2 named. That is the shape the fix
was written for, and the long tail is the same archives refused for a different
reason first.

Class B's remaining 1,061,765 verdicts are change 0580's finding-1 relaxation,
untouched: eleven other local-versus-central consistency checks that no longer run
for a record the caller is not reading. **This change does not close finding 1 and
does not claim to.** Change 0582's recommendation — that change 0580 be
re-approved on its corrected delta table rather than on the single-row one — still
stands, now against a table 32,472 verdicts smaller.

Class B and class E remain **zero on every real archive and every seed**, on both
builds.

## Limitations

**This fix narrows the parser-differential class; it does not close it.** A
complete closure is equivalent to the archive-wide proof change 0580 removed, and
the reason is structural: a reader that follows local headers walks records
sequentially from offset 0, so whether it agrees with us about *any* member's
boundaries depends on *every* earlier record's local header. A target-scoped proof
reads a bounded set of neighbours by construction, so it can only narrow the
disagreement. Three residual shapes were built and measured rather than argued;
`residual_witness.py` generates all three and `residual-verdicts.txt` carries
their verdicts.

| residual | construction | pre-change | 0580 + 0583 |
| --- | --- | --- | --- |
| `residual-prune-distant-csize.zip` (200,188 bytes) | predecessor at 0, central `compressed_size` 16, **local 300,000**; target at 200,000, past the zero-I/O bracket end of 131,140, so the predecessor's local header is **never read** and the maximum never runs | REFUSE `InvalidSize { expected: 16, actual: 300000 }` | **accept, 16 bytes** |
| `residual-zip64-sentinel-csize.zip` (388 bytes) | predecessor's local `compressed_size` is the ZIP64 sentinel with a local ZIP64 extra field declaring 300,000; the probe does not read the variable region | REFUSE `local ZIP64 size framing requires version 45 or newer` | **accept, 16 bytes** |
| `residual-descriptor-flag-disagreement.zip` (388 bytes) | **central** record declares a data descriptor, **local** header does not, local `compressed_size` 100,000, with a valid descriptor exactly where the central length puts it; a reader following local headers sees bit 3 clear and uses the inflated size, but the bound is a `Descriptor` and must not move its resolve offset | REFUSE `strict local and central flags differ` | **accept, 16 bytes** |

All three are residuals of **change 0580**, not regressions of change 0583: the
pre-change build refuses every member of all three, and change 0583 leaves each
exactly where change 0580 put it. Closing the first needs one 30-byte read per
predecessor — the archive-wide cost change 0580 removed, which would falsify its
measured result. Closing the second needs a second, variable-length read of the
neighbour's ZIP64 extra field, up to 131,070 bytes. Closing the third needs a
bound that can refuse *without* resolving, which the current `LocalSpanBound`
cannot express, or a new refusal identity for a framing disagreement on a record
nobody is reading — both larger than this defect fix and both belonging to a
re-approval of change 0580's delta. None is free, and none is attempted here.
They are named so the residual is on the record.

**The class-C defect this change nearly shipped is itself a limitation of the
gate.** The differential found zero accept → refuse divergences over 1,924,524
member verdicts on the intermediate implementation that carried it, because the
shape needs a local/central bit-3 disagreement *plus* a valid descriptor displaced
by fewer than 12 bytes, and change 0582's operator list does not generate that
combination. It was found by reading the code, not by running the corpus. A
deterministic corpus proves the absence only of the shapes someone thought to
write.

**The `parse_zip` fuzz gate remains outstanding.** `cargo-fuzz` is not installed
on this host and no nightly toolchain is present, exactly as change 0582 recorded.
Nothing here discharges it.

**Change 0582's own bounds carry over unchanged**, because this re-run uses that
corpus: the mutation family is weighted toward the first six records of each base
archive; no ZIP64 archive larger than 4 GiB was exercised; and absence of a class
C or D divergence over 1,924,524 member verdicts is evidence, not a proof of
soundness.

**Concurrency was not re-exercised.** The harness is single-threaded by
construction. The order-independence and memo-stability oracles that this change
had to satisfy are single-threaded properties of one reader; the single-flight
contract and the re-entrancy error are unchanged and untested by this run.

**No performance measurement was taken.** The census diff shows change 0580's read
and byte counts did not move, which is a statement about those counts and nothing
else. No timing, allocation, cold-cache, range-source or cross-platform capture
was run for this change.

## Retained evidence

Under [`results/change-0583/`](results/change-0583/): `run.sh`, the whole gate
end to end; `prefix_archive.py`, which reconstructs change 0580's `archive.rs`
by undoing this change's two hunks and asserts the result's sha256 against the
file change 0582 measured, so the pre-fix tree is that code rather than an
approximation of it; `classification-summary.json`; `delta.py` and
`delta-summary.json`; `crafted-verdicts.txt` (1,816 divergent crafted rows across
all three builds); `residual_witness.py` and `residual-verdicts.txt`; `mutate.py`
and `mutation-matrix.txt`; `error-vocabulary.txt`; `prechange-tests.txt`; and
`test-summary.txt`. The packet's own index is its `README.md`.

The two full differential reports are 238,685,636 and 235,006,954 bytes and are
**not** retained. The census output is not retained either: it is byte-identical
to change 0580's `results/change-0580/census-after.txt`, which `run.sh` diffs
against.
