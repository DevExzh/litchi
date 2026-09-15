# 0582: a deterministic differential for change 0580's ZIP strict-layout scope — no soundness counterexample, and a shipped delta 30× the approved table

Status: retained; **change 0580 is recommended for a hold, not a revert**. See
"Recommendation".

`performance_claim: none`. This is a correctness and safety gate, not a
performance change. No timing, allocation, read-count or byte-count result is
measured or claimed here, and nothing under `crates/` was modified.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

Change [0580](0580-zip-target-scoped-strict-layout.md) narrows a ZIP refusal: it
replaces an archive-wide strict-layout proof with one scoped to the member being
read. It ran every gate `docs/GOAL.md` names except one. Its own words:

> The `parse_zip` fuzz target was **not** run: `cargo-fuzz` is not installed in
> this environment and no nightly toolchain is present.

`docs/GOAL.md` lists "existing fuzz targets" among the required gates, and rule
12 forbids weakening malicious-input defenses. For a change that makes a refusal
*less* strict, that is the gate that matters most. This record closes as much of
that hole as a stable toolchain can, and says plainly what it does not close.

The absence was verified rather than assumed:

| claim | verified |
| --- | --- |
| no `cargo-fuzz` on this host | `which cargo-fuzz` → not found |
| no nightly toolchain | `rustup toolchain list` → stable 1.95.0 plus pinned 1.56/1.65/1.89/1.98 only |
| `results/change-0457/fuzz/artifacts/post-run/zip/parse_zip.gz` is not a corpus | `file` → "gzip compressed data … original size 23226512"; it is change 0457's compiled 23 MB fuzz binary, built from that change's `soapberry-zip`, 125 change records ago |
| `corpus-inventory.json` lists content that was not retained | 98 entries, 3,023,111 bytes, each a `{path: <sha1>, bytes, sha256}` row with no file beside it |
| retained ZIP seeds | **7**: the six under `results/change-0457/fuzz/seeds/zip/`, plus `seeds/native/native-small.odp`, which is itself a ZIP container |

## What was built

A plain, deterministic, reproducible differential harness. No `cargo-fuzz`, no
nightly, no new dependency in any crate.

* `results/change-0582/build_corpus.py` — the corpus generator. Seeded once from
  the constant `RNG_SEED = 0x05820580`, drawn from in a fixed order over
  sorted paths, so the whole corpus regenerates byte for byte.
* `results/change-0582/strict_scope_differential.rs` — the harness, built as an
  `examples/` target of `soapberry-zip` inside each extraction so it links the
  crate under test without adding a dependency anywhere.
* `results/change-0582/classify.py` — the comparator.
* `results/change-0582/strict_scope_coverage.rs` — a separate coverage probe.
* `results/change-0582/run.sh` — the whole thing, end to end.

### The two builds

Both are `git archive` extractions of the committed revision `93a610ded`, each
with its own `CARGO_TARGET_DIR` under the session scratchpad. Nothing was built
against the shared working tree.

| tree | contents |
| --- | --- |
| **before** | `93a610ded`, unmodified |
| **after** | `93a610ded` with change 0580's two files overlaid |

The two overlaid files were snapshotted at the start of the run and re-checked
at the end, so a concurrent edit by another agent could not have reached either
build:

```
df9ed280e21112cbd99fd5bb373f50b676c7d592e0b15117ecff6e89bc673b27  crates/soapberry-zip/src/archive.rs
79845d6ca62e0817c7e35879edbb49080a378727c18fb99d9337c807ff486892  crates/soapberry-zip/src/office.rs
```

Host: Linux 7.0.0-1012-aws, AMD EPYC 9R45, 32 cores; rustc 1.95.0, release
profile, `-j 8`, one cargo process at a time.

### The corpus

22,875 inputs, 334,625,602 bytes. `results/change-0582/corpus-fingerprint.txt`
carries the SHA-256 of every file and one hash over the whole sorted list.

| family | inputs | what |
| --- | ---: | --- |
| `seeds` | 7 | the ZIP-format seeds change 0457 retained |
| `testdata` | 516 | every real ZIP container under `test-data/`. Change 0573's census found 533; 17 are byte-identical duplicates of another fixture and dedupe to 516 distinct archives |
| `crafted` | 45 | hand-built archives aimed at the strict-layout proof |
| `mutations` | 22,307 | deterministic mutations of 92 bases — the 7 seeds, the 45 crafted archives, and 40 real archives under 24 KiB selected by a fixed stride over the sorted path list |

Mutation operators, all deterministic: truncation at every structurally
interesting offset (each EOCD, central-directory and local-header boundary, each
name/extra/payload boundary, plus fixed fractions); single-field corruption of
local headers (`signature`, `flags`, `method`, `crc`, `compressed_size`,
`uncompressed_size`, `file_name_length`, `extra_field_length`); single-field
corruption of central headers (the same plus `local_header_offset`); every
central record repointed at the first record's local header; duplicated,
reversed, seed-shuffled and mis-counted central directories; EOCD count/size/
offset corruption; and seeded single-bit flips confined to framing regions so
the input still reaches the strict path instead of dying in inflate.

The 45 crafted archives are the load-bearing half. They include change 0575's
**exact** adversarial witness — the generator reproduces it at 4,405 bytes, with
A's 4,096-byte local extra field under a `0xFACE` header, offsets 0 / 2048 /
4163 and the directory at 4230 — plus inflated local extra fields that swallow
one, two or three following records; spans reaching into the central directory
from the target and from a predecessor; a predecessor whose central compressed
size alone reaches the target; descriptor-bearing members (unsigned, signed and
ZIP64) gapless, at the boundary and overlapping; duplicate local-header offsets;
a local header placed inside another record's payload; exact adjacency, one byte
of slack and one byte of overlap; directory records participating in an overlap;
records whose local and central metadata disagree on name, CRC, size, method,
flags, signature and encryption flag; and the residual-window cases below.

### The surface driven — and what was dropped

`exercise_fuzz_target_body` in the harness is a port of the whole of
`crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs`: `ZipArchive::from_slice`
plus entry iteration and path probes, then `exercise_bounded_paths` with
`exercise_borrowed_reader`, `exercise_reader_at_metadata`,
`exercise_preservation_index`, `exercise_replay`, `exercise_fused_precompressed`,
`exercise_precompressed` and `exercise_precompressed_republication`, including
every `assert!` in them — those asserts are the oracle and a failure is recorded
as a panic. Only one thing was changed: `BoundedSink`, the limit constants and
the control flow are verbatim, but the `fuzz_target!` wrapper becomes a `main`
that iterates a directory and catches unwinds per input.

**What was dropped, and why.** libFuzzer itself: coverage feedback, the mutation
engine, `-rss_limit`/`-timeout` crash detection, `AddressSanitizer`, and the
`libfuzzer-sys` arbitrary-input plumbing. There is no coverage-guided search
here at all — the corpus is fixed. That is the whole of the difference and it is
the reason the Limitations section says this does not replace the real gate.

**What was added, because the real fuzz target barely reaches the changed code.**
Reading `parse_zip.rs`, only two call sites reach the strict-layout proof
(`read_entry_precompressed_with_progress` and
`read_entry_precompressed_and_decoded_with_progress`), and both `break` after
the first member. `ArchiveReader::read`, `read_stored_borrowed` and
`preservation_index` do not reach it at all: `read` goes through
`archive.get_entry`, and the other two keep their own archive-wide proofs. So a
`parse_zip` run exercises the proof roughly once per input.

The harness therefore adds a full sweep of **every** member through **all four**
public entry points that reach `strict_layout_for`:

| entry point | reader |
| --- | --- |
| `ArchiveReader::read_to` | slice-backed |
| `IndexedArchive::read_entry_to` | source-backed |
| `IndexedArchive::with_verified_entry_reader` | source-backed |
| `IndexedArchive::read_entry_precompressed_and_decoded_with_progress` | source-backed |

and records, unchanged, the three archive-wide surfaces 0580 says it does not
touch (`read_stored_borrowed`, `preservation_index_with_limits`, and both opens)
so a regression in them lands in the same diff.

Two limit profiles run over the whole corpus. `fuzz` is `parse_zip.rs`'s own
`ArchiveLimits` verbatim. `wide` exists because those limits (256 files, 64 KiB
of metadata, 1 MiB total) refuse most real Office fixtures at open, which would
make the `testdata` half of the corpus prove nothing about this change.

Two intra-build oracles run on every input, on both builds: the same member set
read in reverse order on a fresh reader must produce the same verdicts, and a
second pass on the same reader must produce the same verdicts. These pin the
order-independence property 0580 claims for its memo.

## How much of this reached the changed code

A big input count that never reaches the changed code is not evidence, so this
was measured rather than argued. A **third**, separately built copy of the after
tree carries two atomic counters — one on `strict_layout_for`, one on
`prove_target_scoped_strict_layout`. It is not one of the differential builds
and no verdict in this record comes from it.

| family | inputs | inputs entering the proof | proof entries | proof builds |
| --- | ---: | ---: | ---: | ---: |
| `seeds` | 7 | 7 | 304 | 228 |
| `testdata` | 516 | 516 | 218,736 | 164,052 |
| `crafted` | 45 | 44 | 7,776 | 6,412 |
| `mutations` | 22,307 | 14,901 | 3,590,520 | 2,997,630 |
| **total** | **22,875** | **15,468** | **3,817,336** | **3,168,322** |

**7,407 of 22,875 inputs (32.4%) never entered the strict-layout proof.** They
are refused earlier — by EOCD discovery, the central-directory parse, or the
archive limits — and prove nothing whatever about this change. The one crafted
archive that does not enter it is `empty-archive.zip`, which has no members. The
figure for the corpus that matters is that **every real archive and every seed
entered**, and two thirds of the mutations did.

## Results

2,692,431 verdicts were compared: 1,924,524 member verdicts across the four
strict-layout entry points and both limit profiles, 603,614 on the archive-wide
surfaces and the two opens, 40,361 slice-level parse records, and 123,932
order-independence and memo-stability assertions.

| outcome | count |
| --- | ---: |
| **panics, either side** | **0** |
| **intra-build oracle failures, either side** | **0** |
| accept → refuse (class C) | **0** |
| accept → accept with different bytes (class D) | **0** |
| refuse → accept, pre-change refusal was the overlap error (class A) | 37,176 |
| refuse → accept, pre-change refusal was something else (class B) | 1,094,221 |
| refuse → refuse with a different typed error identity (class E) | 351,525 |
| identical | 1,209,509 |

Every divergence lands on one of the four strict-layout entry points. The
603,614 archive-wide verdicts (`read_stored_borrowed`,
`preservation_index_with_limits`), the 40,361 slice-level parse records and both
opens are **byte-for-byte identical** on both builds, which is the control that
0580 did not disturb the paths it says it left alone.

### The real-corpus control

| family | strict member verdicts, before | after |
| --- | --- | --- |
| `testdata`, `fuzz` profile | 54,023 accept, 0 refuse | 54,023 accept, 0 refuse |
| `testdata`, `wide` profile | 55,179 accept, 0 refuse | 55,179 accept, 0 refuse |
| `seeds`, both profiles | 74 accept, 0 refuse | 74 accept, 0 refuse |

**Not one verdict changes on any real container in the repository.** All 516
distinct archives, 109,202 member verdicts across the two profiles, all
accepted, identically. (The remaining rows on those families — 69 and 97 — are
the harness's own eight-event progress-callback cap on the precompressed path,
and are identical on both sides.)

That control also establishes a fact used below: because the **pre-change**
archive-wide proof accepted every member of every real archive, every record of
every real archive has local and central metadata that agree.

### The approved witness reproduces exactly

| target | before | after |
| --- | --- | --- |
| `A.bin` (overlaps B) | REFUSE, "strict streaming refuses overlapping ZIP local spans" | REFUSE, same identity |
| `B.bin` (overlapped by A) | REFUSE, same identity | REFUSE, same identity |
| `C.bin` (overlaps nothing) | REFUSE, same identity | **accept, 32 bytes** |

On all four entry points, both readers, both limit profiles. This is the one row
the project owner approved, and it is exactly what the after build does.

### The residual window holds at its extreme

Change 0580 re-derived change 0575's predecessor bracket, doubling the residual
window from 65,559 to 131,094 bytes on the grounds that a record nobody validates
may declare a local `file_name_length` its central record does not carry. Two
crafted archives price that at the boundary:

| witness | construction | after |
| --- | --- | --- |
| `predecessor-inflated-local-name` | change 0580's own attack, rebuilt: 9-byte central name, 32-byte payload, local `file_name_length = 100` and `extra_field_length = 65535`, declared span end 65,697; target at 65,640, which is **past** 0575's 65,630 bracket end and **inside** the declared span | **REFUSE**, overlap |
| `window-max-residual-reaches` | local `file_name_length = 65535` *and* `extra_field_length = 65535`, span end 131,132 — the largest span a `u16` pair can describe; target at 131,100 | **REFUSE**, overlap |
| `window-max-residual-clear` | the same predecessor, target at 131,140, eight bytes past the span end | accept |
| `neighbour-local-extra-inflated` | local `extra_field_length` inflated instead of the name | **REFUSE**, overlap |
| `predecessor-csize-reaches` | central compressed size alone reaches the target, so no local header need be consulted | **REFUSE**, overlap |
| `neighbour-local-csize-understates` | local size 16, central size 4096 — the bound takes the larger | **REFUSE**, overlap |

The correction is sound at the extreme it was derived for, and the boundary is
exact to the byte.

### Two claims of 0580, checked against the corpus

**The two removed error identities were dead code.** `"strict streaming local
span extends into the central directory"` and `"strict layout proof publication
failed"` appear **zero** times in the before report over 22,875 inputs and
3.8 million proof entries. That corroborates 0575's archaeology rather than
proving unreachability, but it is the strongest evidence a black-box run can
give.

**The duplicate-offset refusal was unreachable and is now reachable.**
`"strict streaming refuses duplicate ZIP local spans"` appears **zero** times
before and **44,328** times after. 0575 found the branch had no test and was in
fact unreachable; 0580 made it reachable, and the corpus confirms it. This is a
*strengthening*, and it is the one string added to the error vocabulary. No
error identity disappeared: diffing the full set of `InvalidInput` messages that
either build emits over the whole corpus — 39 distinct identities before, 40
after — yields exactly one line, and it is an addition.

## Finding 1 — the approved table is 3.3% of the delta that shipped

**This is the headline finding and it is reported rather than rationalised.**

The approved behavioural delta is one row: a member that overlaps nothing
becomes readable in an archive where two *other* members overlap each other.
That is class A: **37,176 member verdicts over 480 inputs**.

The delta actually shipped is class A **plus** class B: **1,094,221 member
verdicts over 10,349 inputs** in which the pre-change refusal was not an overlap
at all, but some other per-record check on a record the caller never reads:

| pre-change refusal that no longer prevents reading another member | verdicts |
| --- | ---: |
| `InvalidSize { expected, actual }` — local and central sizes disagree | 353,359 |
| `"strict local and central flags differ"` | 160,188 |
| `InvalidChecksum { expected, actual }` — local and central CRC disagree | 127,350 |
| `"strict local and central compression methods differ"` | 121,594 |
| `"strict local and central names differ"` | 84,686 |
| `"local ZIP64 size framing requires version 45 or newer"` | 63,864 |
| `"stored local and central flags differ"` | 53,404 |
| `Eof` — the record's declared framing runs off the source | 45,210 |
| `"stored local and central compression methods differ"` | 40,538 |
| `"stored local and central names differ"` | 28,234 |
| `InvalidSignature { expected, actual }` — the record has no local header signature | 15,026 |
| `"borrowed access refuses nonzero disk-start metadata"` | 768 |

**Twelve distinct refusal identities, not one.** 1,094,221 verdicts against the
approved class's 37,176: the shipped delta is 1,131,397 member verdicts, of
which the approved row accounts for 3.3%, and the whole is 30× the part that was
put to the owner.

Change 0580's prose does disclose this. Its "semantic difference, in bytes"
section lists "its local and central compression method, flags, name, sizes and
CRC; its local ZIP64 size framing; its data descriptor's contents" among the
things no longer checked for an untouched record. So nothing here contradicts
the record. What this measurement adds is the **magnitude and the shape**: an
approval given on a table with one row, derived from a 4,405-byte witness
demonstrating overlap, was in practice an approval of "every local-versus-central
consistency check on every record the caller does not read is dropped", which is
a materially different sentence.

Class B is zero on every real archive and every seed. It is reachable only on
archives that are already malformed.

## Finding 2 — a local-versus-central size disagreement now smuggles a readable member

This is the sub-class of finding 1 with an identifiable adversarial shape, and
it is worth naming separately because it is a parser-differential rather than a
relaxation.

`crafted/neighbour-local-csize-smuggles.zip` (388 bytes) has two Store members:

```
pred.bin     local offset 0    central compressed_size 16   LOCAL compressed_size 100000
target.bin   local offset 200  central compressed_size 16   LOCAL compressed_size 16
```

A reader that trusts local headers — which is what a streaming ZIP reader must
do, since a data-descriptor archive cannot be read any other way — places
`pred.bin`'s payload at `[38, 100038)`, which covers `target.bin`'s entire local
record. A reader that trusts the central directory places it at `[38, 54)`,
which does not.

| target | before | after |
| --- | --- | --- |
| `pred.bin` | REFUSE `InvalidSize { expected: 16, actual: 100000 }` | REFUSE, same identity |
| `target.bin` | REFUSE `InvalidSize { expected: 16, actual: 100000 }` | **accept, 16 bytes** |

`crafted/neighbour-local-csize-smuggles-deflate.zip` shows the same with Deflate
members: **480 decoded bytes of `target.txt` are returned** where the pre-change
build refused. `neighbour-local-usize-smuggles` flips the same verdict through the
uncompressed-size field, though only the compressed-size form moves where a
local-header-trusting reader would place the payload; the uncompressed form is
an ordinary finding-1 instance.

The mechanism is in `local_span_bound_from_fixed` in `archive.rs`. It reads the
predecessor's 30-byte fixed local header, takes the **variable-region length**
from that header — so an inflated local name or extra field *is* caught, which
is why finding 1's window tests refuse — but takes the **payload length** from
the central record and never compares it against the local header's own
`compressed_size` field — a field `ZipLocalFileHeaderFixed::parse` has already
decoded at the top of the same function, and which the descriptor branch just
below reads for the ZIP64 sentinel:

```rust
let payload_end = entry.local_header_offset
    .checked_add(ZipLocalFileHeaderFixed::SIZE as u64)
    .and_then(|offset| offset.checked_add(variable_length))
    .and_then(|offset| offset.checked_add(entry.compressed_size))   // central, always
```

Before 0580 this was unreachable, because the archive-wide proof refused any
record whose local and central sizes disagreed before any member could be read.

**What limits it.** The two archive-wide proofs 0580 deliberately left alone
still refuse the same archive: `read_stored_borrowed` reports
`InvalidSize { expected: 16, actual: 100000 }` for *both* members, and
`preservation_index_with_limits` reports
`UnsupportedPreservation { reason: "invalid local member framing" }`. And every
byte returned for `target.bin` is `target.bin`'s own, fully validated against
its own local header, sizes and CRC — there is no confusion *inside*
soapberry-zip, only between soapberry-zip and a local-header-trusting reader.

**What does not limit it.** The narrowed paths are not internal.
`crates/litchi-opc/src/source_backed.rs` reads OOXML parts through
`read_entry_to`, `with_verified_entry_reader`,
`with_verified_entry_reader_with_accounting`,
`read_entry_precompressed_with_progress` and
`read_entry_precompressed_and_decoded_with_progress` — all five reach
`strict_layout_for`. So the narrowing is reachable from the DOCX, XLSX and PPTX
facades, which is the active priority. (`crates/litchi-odf-common/src/core/package.rs`
reaches it through `with_verified_entry_reader` too; ODF is deferred and was not
otherwise examined.)

**A remedy that cannot widen the delta.** Taking the payload length as
`max(local_compressed_size, central_compressed_size)` in
`local_span_bound_from_fixed`, when the record declares no data descriptor and
the local size is not the `u32::MAX` ZIP64 sentinel, closes this at **zero
additional I/O** — the field is already parsed from a header already read. It
can only ever refuse more, never accept more, so it cannot widen the approved
delta. The real-corpus control above says it is a no-op on every archive in
`test-data`: the pre-change archive-wide proof accepted all 516 of them, which
required local and central sizes to agree for every record. Change 0580 owns
that decision; this record only prices it.

## Finding 3 — 351,525 typed-error identity changes on members that stay refused

Class E. Every one of these is a refusal on both sides, so nothing becomes
readable, but `docs/GOAL.md` rule 12 names typed errors among the things to
preserve and the drift is large: **130 distinct before → after identity pairs**.

| before → after | verdicts |
| --- | ---: |
| `Eof` → overlap | 59,604 |
| `InvalidSize` → `InvalidSignature` | 41,512 |
| `InvalidSignature` → `Eof` | 12,586 |
| `InvalidSignature` → overlap | 9,670 |
| `InvalidChecksum` → overlap | 1,480 |
| `Eof` → duplicate local spans | 164 |
| 124 further pairs | 226,509 |

The mechanism is not a defect: the archive-wide loop reported whichever record
failed first in local-header order, so every member of a malformed archive
inherited one stranger's error; the target-scoped proof reports the target's own
failure, or the overlap that actually covers it. The after identity is, if
anything, the more informative one. It is recorded because it is a contract
change nobody has stated, it is large, and a caller that matches on
`ErrorKind` will see different values for the same bytes.

Class E is zero on every real archive and every seed.

## What this does and does not cover, versus the real fuzz target

**Covered, and more thoroughly than `parse_zip` covers it.** Every public entry
point that reaches `strict_layout_for` — four of them — on every member of every
input, in forward and reverse order, on fresh and reused readers, under two
limit profiles, with a typed error identity and a payload fingerprint recorded
for each. `parse_zip` reaches that proof through one API, for one member per
input.

**Covered, at parity.** `ZipArchive::from_slice`, EOCD and ZIP64 discovery,
central-directory iteration and path normalization, the borrowed reader, the
`ReaderAt` metadata surface, the preservation index, the precompressed capture
and republication path with its short-write and partial-sink sub-cases, and
`write_replacing_with_replay` with its drift and partial-sink sub-cases —
all ported from `parse_zip.rs` with their assertions intact.

**Not covered.** Coverage-guided input generation; libFuzzer's mutation engine
and corpus minimisation; `AddressSanitizer`, `MemorySanitizer` and
`LeakSanitizer`; RSS and timeout crash detection; and `-max_len` boundary
search. This corpus is fixed and its mutations are hand-specified operators, not
a search.

**Not covered, and orthogonal to this change.** Concurrency. The single-flight
contract, the re-entrancy error and the two existing concurrency tests are
unchanged by 0580 and were not re-exercised here; the harness is single-threaded
by construction so its results are byte-reproducible.

## Limitations

**This is a deterministic differential harness. It is not a replacement for
coverage-guided fuzzing.** It cannot find an input nobody thought to write. Its
22,875 inputs come from seven retained seeds, 516 real archives, 45 archives
written by hand against a reading of the changed code, and 22,307 mutations from
a fixed operator list. A coverage-guided fuzzer explores the space between those
points, and this does not. **The `parse_zip` gate remains outstanding on this
host for want of `cargo-fuzz` and a nightly toolchain**, and this record does not
discharge it — it reduces how much of it is unknown.

Three further bounds on what the numbers mean:

* **Absence of a class C or D divergence is not a proof of soundness.** It says
  that over 1,924,524 member verdicts, no input made the after build accept
  bytes the before build refused *for the member being read*, and none made it
  return different bytes. A soundness defect reachable only by an archive shape
  outside this corpus would not appear.
* **The mutation family is heavily weighted toward the first six records** of
  each base archive, because the operators index records `[..6]` to bound the
  corpus. A defect that needs a malformed record deep in a large central
  directory is under-sampled.
* **No ZIP64 archive larger than 4 GiB was exercised**, and no genuinely
  ZIP64-framed EOCD appeared in `test-data` (change 0573's census notes the
  same). The ZIP64 paths reached here are the descriptor and sentinel forms in
  the crafted family, not a true ZIP64 container.

Nothing here measures time, allocation, read counts or bytes; change 0580 owns
those figures and none of them were re-derived.

## Recommendation

**Hold change 0580 for a re-approval on its true delta. Do not revert it.**

The implementation survived the strongest black-box test a stable toolchain can
build: zero panics over 22,875 inputs and 3.8 million proof entries, zero
accept-to-refuse, zero payload changes, byte-identical behaviour on every real
container in the repository, an exact reproduction of the approved witness, a
residual window that is sound to the byte at its extreme, and a refusal branch
that went from unreachable to reachable. **No counterexample to the soundness of
the target-scoped proof was found.** Reverting would throw that away and
reinstate an archive-wide cost that change 0572 priced at 264 of the 354
requests an XLSX cell read issues.

What is not yet settled is the approval. The owner approved one row on one
witness. The change delivers twelve refusal identities over 1.09 million member
verdicts, including a local-versus-central size disagreement that lets a member
be read out of a region a local-header-trusting reader assigns to a different
record, reachable from `litchi-opc` and therefore from the OOXML facades. Two
things are worth putting in front of the owner before this lands:

1. the corrected delta table above, in place of the single-row one; and
2. finding 2's remedy — `max(local, central)` compressed size in the neighbour
   bound — which costs no I/O, cannot widen the delta, and is a no-op on every
   archive in `test-data`.

Change 0575's falsification criterion 1 — "a rationale for the archive-wide
scope exists outside the repository" — is still the standing rule. This record
adds a second: if the archive-wide local-versus-central consistency check was
load-bearing for any caller, finding 1 names the twelve identities that stopped
running.

## Retained evidence

Under `results/change-0582/`:

| file | what |
| --- | --- |
| `build_corpus.py` | the corpus generator; `RNG_SEED = 0x05820580` |
| `strict_scope_differential.rs` | the harness |
| `strict_scope_coverage.rs` | the coverage probe |
| `classify.py` | the comparator |
| `run.sh` | both builds, the corpus, both runs and the classification, end to end |
| `corpus-fingerprint.txt` | family counts, per-file SHA-256 for the crafted family, and one hash over the whole sorted manifest |
| `classification-summary.json` | every count in this record |
| `coverage-summary.txt` | the coverage probe's per-family table |
| `crafted-verdicts.txt` | the full before/after verdict table for all 45 crafted witnesses |
| `README.md` | the packet's own index |

The two full reports are 238,685,636 and 233,872,704 bytes and are **not** retained; `run.sh`
regenerates them, and `corpus-fingerprint.txt` plus
`classification-summary.json` are what this record cites.

## Disposition: finding 2 is fixed as change 0583

Recorded after the fact; nothing above has been rewritten.

**Finding 2 was fixed** as change
[0583](0583-zip-local-size-span-bound.md), by the remedy this record proposed in
"A remedy that cannot widen the delta": `local_span_bound_from_fixed` now
reserves the larger of the local header's and the central record's
`compressed_size` for a neighbour's payload. One new function in
`crates/soapberry-zip/src/archive.rs`, no extra I/O, six new tests.

Three refinements to what this record proposed, each measured rather than
argued:

* The remedy's own condition — "when the record declares no data descriptor and
  the local size is not the `u32::MAX` ZIP64 sentinel" — is load-bearing, not
  merely conservative, and for a sharper reason than conservatism. A
  `LocalSpanBound::Descriptor`'s `payload_end` is the offset the descriptor is
  *read at*, not only a threshold, so moving it does not enlarge the bound: it
  points the read somewhere else, where a descriptor can match that did not
  match before. An implementation that applied the maximum there produced a
  genuine **class-C divergence** — a member refused before and readable after —
  and, separately, **244 intra-build order-independence oracle failures**,
  because a descriptor-bearing record validates with disagreeing sizes and has
  its exact span end memoised as its neighbour bound. Change 0583 confines the
  maximum to the exact branch. Both defects were found before anything landed,
  one by this harness and one by code review.
* **This corpus did not find the class-C defect.** Zero accept → refuse
  divergences over 1,924,524 member verdicts on the build that carried it,
  because the shape needs a local/central bit-3 disagreement *plus* a valid
  descriptor displaced by fewer than 12 bytes, which the operator list in "The
  corpus" above does not generate. That is a concrete instance of the
  Limitations section's "it cannot find an input nobody thought to write".
* This record's reachability claim holds, with one correction: the narrowing is
  reachable from the DOCX, XLSX and PPTX facades, but **not** from
  `PartView::data()`, the ordinary part read, which goes through
  `read_entry_with_accounting` and never enters the strict-layout proof.
  `strict_layout_for` also lives in `office.rs`, not `archive.rs`.

**The re-run**, using this record's harness, corpus (verified byte-identical to
`corpus-fingerprint.txt`), classifier and limit profiles, against a build
carrying change 0580 + change 0583:

| outcome | this record | after change 0583 |
| --- | ---: | ---: |
| panics, either side | 0 | **0** |
| intra-build oracle failures | 0 | **0** |
| accept → refuse (class C) | 0 | **0** |
| accept → accept, different bytes (class D) | 0 | **0** |
| class A | 37,176 | 37,160 |
| **class B** | **1,094,221** | **1,061,765** |
| class E | 351,525 | 383,981 |

**32,472 member verdicts that change 0580 made readable are refused again**, and
**zero** verdicts that change 0580 refused become readable — measured by pairing
all three builds verdict by verdict. The finding-2 witnesses
`neighbour-local-csize-smuggles.zip` and `-deflate.zip` refuse on both sides
again, on all four entry points and both limit profiles, with the overlap
identity; the 480 decoded bytes of `target.txt` are no longer returned.
`neighbour-local-usize-smuggles` is deliberately unchanged — this record's
classification of it as an ordinary finding-1 relaxation is correct, because
`uncompressed_size` never describes an on-disk region.

**Finding 1 is not closed and finding 3 is not closed.** Class B still carries
1,061,765 verdicts across eleven other local-versus-central consistency checks,
and the class E identity drift grows. This record's recommendation — that change
0580 be re-approved on its corrected delta table rather than the single-row one —
still stands, now against a table 32,472 verdicts smaller. Change 0583's own
Limitations section names three residual shapes it does not close and prices all
three with retained witnesses.

The `parse_zip` gate remains outstanding for the same reason as here.
