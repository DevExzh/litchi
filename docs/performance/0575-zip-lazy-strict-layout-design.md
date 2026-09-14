# 0575: design for a target-scoped ZIP strict-layout proof

Status: design only. No production change and `performance_claim: none`. This
record freezes a design, an information-theoretic impossibility result, an
adversarial witness and a numeric prediction; it measures no latency and claims
no improvement.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

`docs/GOAL.md`'s definition of done requires that "selective reads perform work
proportional to mandatory metadata plus accessed content rather than total
uncompressed document size where the format permits". Reading one member of a
source-backed OOXML package violates the spirit of that clause in a way the
clause's own wording does not quite name: the work is proportional not to the
uncompressed size but to the archive's **member count**, and it is paid in
positional reads before the first payload byte.

On `ConditionalFormattingSamples.xlsx` (132 central records, measured below) the
first read of any member issues **264 positional reads** that touch 10,740 bytes
of local-header framing. None of them belong to the member being read.

Change [0567](0567-ooxml-single-index-per-open.md) found this and deliberately
declined to act:

> The proof re-reads local-header framing the central directory already
> describes, which is the class change 0561 identified, but it establishes a
> whole-archive non-overlap property rather than a per-entry one, so it cannot be
> made per-entry without weakening it. This is recorded, not addressed: it is a
> correctness-sensitive security check and a 4.45 microsecond warm-cache saving
> does not justify touching it.

0567 also measured it, at a very consistent 319-345 ns per member: **4.45 µs** on
`sheet-names.xlsx` (13 members), 15.26 µs on `Tables.xlsx` (46) and **42.94 µs**
on `ConditionalFormattingSamples.xlsx` (132), or 6.2% of the open-to-first-cell
interval on the smallest of the three. The 4.45 µs in the passage above is the
13-member figure, which is the weakest case the decision could have been made on.

This record tests both halves of that sentence. The cost half is correct for a
warm local file and is the reason nothing has been done — and 0567's own next
sentence is the licence for this record to reopen it:

> Change 0561's own governing conclusion applies — justify such work against a
> latency-bearing source, not a warm-cache count.

The property half is **half right**: the whole-archive property genuinely cannot
be reduced to a per-entry one without weakening it, and §"The lower bound" below
proves that formally. But a third property sits between the two: it is
order-independent, it converges to the whole-archive property, and it is what
actually protects a read.

## What the proof establishes today, precisely

Two implementations exist and they are structurally identical:

| | owner | source | cost |
| --- | --- | --- | --- |
| `ArchiveReader<'data>` | `office.rs:2506` | `&'data [u8]` | **zero I/O**, O(n) CPU |
| `IndexedArchive<R: ReaderAt>` | `office.rs:3582` | positional | **2 positional reads per central record** |

`LazyArchiveReader` wraps `ArchiveReader`, so the in-memory OPC path pays the CPU
and the source-backed OPC path (`litchi-opc/src/source_backed.rs:5305`,
`:5656`, `pkgreader.rs:307`) pays the I/O.

Both loop over `layout` — every central record, **including directory records** —
sorted by local-header offset, and establish four things:

1. **Per-entry layout.** `validate_strict_entry_layout` →
   `validate_reader_entry_layout_with_name_policy` (`archive.rs:728`) checks that
   the local header's method, flags and name equal the central record's, that
   local and central sizes agree (unless a data descriptor is declared), that the
   local CRC agrees, and computes `data_start_offset`, `data_end_offset` and
   `span_end`, resolving a data descriptor when bit 3 is set.
2. **Chain non-overlap.** `span.local_header_offset < previous_end` is refused.
   Because `span_end >= local_header_offset` always holds, `previous_end` is
   non-decreasing once the check passes, so the adjacent-pair chain is
   equivalent to pairwise non-overlap over all records.
3. **Distinct local-header offsets.** Inserting into `by_local_header` twice is
   refused.
4. **No intrusion into the central directory.** After the loop,
   `previous_end > directory_offset` is refused.

Three refinements matter to the design and none of them is documented:

- **(3) needs no I/O at all.** `StrictEntryLayout::local_header_offset` is copied
  verbatim from the wayfinder (`archive.rs:915`). The duplicate test is therefore
  a pure central-directory property, decidable in one pass over already-resident
  metadata.
- **(4) is dead code.** `validate_reader_entry_layout_with_name_policy` already
  refuses `local_end > central_directory_offset`, `data_end_offset >
  central_directory_offset` and `span_end > central_directory_offset` per entry.
  Every span in `spans` therefore satisfies the post-loop predicate before the
  post-loop check runs. The archive also precomputes a zero-I/O approximation of
  the same fact at index build (`all_local_spans_bounded`, `office.rs:894-897`),
  whose accessor's doc-comment (`office.rs:3503-3507`) already names the house
  pattern: "The result is captured during index construction and is therefore a
  zero-pass metadata fact".
- **(2) is the only invariant that costs I/O**, and it costs all of it.

### The proof is not on every read path

`strict_layout_for` is reached from `read_to`, `read_entry_to`,
`with_verified_entry_reader*` and `capture_precompressed`. It is **not** reached
from `IndexedArchive::read`/`read_entry`, which route through
`IndexedReadSession::read_entry_with_accounting` (`office.rs:964`) and call
`self.archive.archive.get_entry(...)` directly, nor from `ArchiveReader::read`.
`archive.rs:318-319` states the intent for the borrowed variant — "This does not
alter ordinary owned reads" — but no record states it for the strict proof.

The consequence is load-bearing for every argument below: **the archive-wide
non-overlap property is not an invariant of the library today.** It is an
invariant of the sink-based, borrowed and callback read families. The same
`IndexedArchive` will accept an overlapping archive through `read_entry` and
refuse it through `read_entry_to`. Change 0567's description of it as a security
check has to be read against that.

## Contract archaeology

Searched: all 30 files under `docs/adr/`, `SECURITY.md`, every numbered record
under `docs/performance/` and `docs/performance/changes/`, all code comments and
doc-comments on the symbols, the `soapberry-zip` test and fuzz corpora, and
`git log -S` on `build_strict_layout_proof`, `StrictLayoutProof`,
`StrictLayoutCache`, `validate_borrowed_spans` and each of the five error
strings.

**No accepted ADR mentions ZIP overlapping entries, ZIP confusion, ambiguous
archives, duplicate local-header offsets, non-overlapping spans, a strict layout,
or a physical-layout proof.** Every hit on "overlap" in `docs/adr/` is a
different domain concept: patch `ConflictSet` overlap (0003:103), BIFF8 range
overlap (0016:49), merged-cell ranges (0008:1097), table coordinates (0028).
ADR 0006 has a `## Security boundaries` section; the word "zip" does not appear
anywhere in that ADR. `SECURITY.md` is a twenty-line disclosure policy. There is
no threat-model document in the repository.

**Neither introducing commit has a message body.**

| Commit | Date | Subject | Body |
| --- | --- | --- | --- |
| `e3d36d6cd` | 2026-08-31 | `fix(zip): validate borrowed stored entries` | *(empty)* |
| `8c0e04159` | 2026-09-01 | `fix(zip): validate indexed stream layouts` | *(empty)* |

`8c0e04159` introduced `build_strict_layout_proof`, `StrictLayoutProof`,
`StrictLayoutCache`, `strict_layout_for_cached` and all three "strict streaming…"
error strings.

**The change records state a mechanism and no rationale.** Change
[0348](changes/0348-stored-zip-borrow-validation.md) — `Status: correctness
evidence recorded; no performance result` — aligns itself with a *performance*
goal:

> This change aligns with `docs/GOAL.md:398` by retaining borrowed access for
> stored ZIP payloads when the source is an immutable slice whose lifetime is
> available.

and mentions overlap once, as a bare list item with nothing attached: "…
encryption, overlap, duplicate-name safety, and strict refusal of a nonempty
entry with a zero CRC." Change
[0351](changes/0351-indexed-stream-validation.md) — `Status: correctness and
resource hardening only` — describes the scope but never why it is archive-wide:

> Every physical span, including directory spans, participates in the layout
> proof. Prefixes and gaps are allowed; overlaps and central-directory intrusion
> are refused.

**Two comments state a design intent, both narrow.** `office.rs:477-478`: "This
is used only to prove non-overlap before publishing a borrowed slice."
`archive.rs:435-438` is the single place where the *global* framing is deliberate:

> Encryption eligibility is intentionally handled by the target path separately
> so an unrelated encrypted member cannot poison the global non-overlap proof.

**The only security framing in the repository is post-hoc**, written twelve days
after introduction, by the same author, arguing against removal: the 0567 passage
quoted at the top of this record. It asserts "correctness-sensitive security
check" and names no threat, no attack and no differential.

One further comment is worth quoting only to rule it out.
`archive.rs:1841-1846` carries a `# Security Usage` block naming zip bombs and
overlapping entries. `git log -S'used in zip bombs'` returns one commit,
`03d30c00d` of 2025-11-29, "Add custom ZIP implementation forked from rawzip" —
it is inherited upstream documentation attached to the unrelated public
`compressed_data_range()` accessor, nine months older than the proof. It is not a
rationale and must not be cited as one.

### What this changes about the risk calculus

The archive-wide scope **accreted without a stated rationale**. That is not the
same as saying it is worthless — the invariant is real and the code is careful —
but it removes the "deliberate security boundary, do not touch" defence. Three
facts follow:

1. There is no recorded threat to preserve, so no candidate can be measured
   against one. Any argument must be made from first principles here, in this
   record, and will become the contract if accepted.
2. The property is already absent from the materializing read family, so
   describing it as a library invariant is not accurate today.
3. The only *documented* driver in the introducing record is a performance goal
   (`GOAL.md:398`, borrowed access for stored entries), and the overlap proof is
   the admission condition that goal needed — not a security feature that
   happened to cost I/O.

## What the central directory alone can prove

The strict path forces `local_name_len == central_name_len` (a differing local
name is refused outright), and for a record without a data descriptor it forces
local and central compressed sizes to be equal. So for record *i* the span end

```
end_i = offset_i + 30 + name_len_i + extra_len_i + compressed_size_i + descriptor_i
```

has exactly **two** unknowns without I/O: `extra_len_i`, a `u16` the central
directory does not carry a copy of, and `descriptor_i`, bounded by the existing
constant `MAX_DATA_DESCRIPTOR_SIZE = 24` (`archive.rs:61`). That gives a
zero-I/O bracket

```
min_end_i = offset_i + 30 + name_len_i + compressed_size_i
max_end_i = min_end_i + 65535 + 24        // the residual window is 65,559 bytes
```

and three verdicts per adjacent pair, all free:

- `max_end_i <= offset_{i+1}` → non-overlap **proven**, no read needed;
- `min_end_i > offset_{i+1}` → overlap **refuted**, no read needed;
- otherwise → **undecided**; only record *i*'s own local header can settle it.

### Measured: the zero-I/O screen proves nothing on real archives

`results/change-0575/zip_layout_census.py` reads the fixture bytes with `struct`
and reports the actual local extra-field lengths.

| Fixture | records | bytes | local extra lengths | adjacent pairs proven | refuted | undecided |
| --- | ---: | ---: | --- | ---: | ---: | ---: |
| `sheet-names.xlsx` | 13 | 9,425 | {0, 264, 520} | **0** | 0 | 12 of 12 |
| `ConditionalFormattingSamples.xlsx` | 132 | 654,688 | {0, 40, 264, 520} | **0** | 0 | 131 of 131 |
| `shapes.pptx` | 48 | 68,822 | {0, 264, 520} | **0** | 0 | 47 of 47 |

No pair is decided in either direction on any fixture. The reason is arithmetic
rather than incidental: OOXML producers pack members contiguously — the measured
gap between one record's exact span end and the next record's offset is **0 on
every pair of all three fixtures** — so the inter-record stride is the member's
own size, typically a few hundred bytes, against a 65,559-byte residual window.
A zero-I/O screen can prune a predecessor only when the bytes between them exceed
64 KiB, which in these packages happens only across the embedded PNGs.

**Candidate (a), a pure zero-I/O pre-screen, therefore proves no non-overlap
whatever on any real OOXML package.** It can still refute, and §"Which existing
tests survive" shows that is not nothing, but it cannot be sold as preserving
invariant (2).

### The lower bound: O(1) is not achievable for invariant (2)

Fix any algorithm that decides invariant (2) for an *n*-record archive. Choose
any *k* and construct two archives identical in every byte except record *k*'s
local extra-field length, chosen so that one archive's spans are pairwise
disjoint and the other's are not. Their central directories are byte-identical,
because the central directory carries a record's *central* extra-field length and
never its local one. Any algorithm that returns different verdicts for the two
must have read a byte in `[offset_k + 28, end_k)`. Since *k* was arbitrary, a
correct decider must be able to read a byte from every record's local header, so
its worst case is **Ω(n) positional reads**.

The task's framing — make a single-member read cost O(1) reads while preserving
(2)–(4) — is therefore unsatisfiable. (3) and (4) are free; (2) is not reducible.
What remains is to choose the strongest property that *is* O(1)-ish, and to name
the difference honestly.

### The property that is worth choosing

For record *t*, define **target disjointness**: no other central record's
declared local span intersects *t*'s span.

- **Successors** (`offset_j > offset_t`) cost **zero I/O**. Once *t*'s own header
  is read, `end_t` is exact, and `offset_j >= end_t` is a comparison against
  resident central metadata. A successor's own extent is irrelevant — it cannot
  reach backwards.
- **Predecessors** cost one read each, but only those in the residual window:
  `min_end_i <= offset_t < max_end_i`. Everything with `max_end_i <= offset_t` is
  pruned free; everything with `min_end_i > offset_t` is refused free.
- Each such predecessor needs only its **30-byte fixed local header**, because
  `name_len + extra_len` plus the central `compressed_size` plus
  `MAX_DATA_DESCRIPTOR_SIZE` bounds its span from above. No second read for the
  variable part, no descriptor read, no name or CRC comparison for a record
  nobody asked to read.

Three properties make this the right choice rather than a compromise:

1. **It is order-independent.** `read(X)` is a pure function of the archive bytes
   and `X`. No memo can change an outcome; a memo only avoids repeating work.
2. **It converges to invariant (2) exactly.** For adjacent pair *(i, i+1)*,
   either the pair is decided free, or `i` is in `i+1`'s residual set. So reading
   every member establishes every adjacent-pair check, which §"What the proof
   establishes" showed is equivalent to pairwise non-overlap.
3. **It is never more expensive than today.** No record's header is read more
   than once, so the total over all members is at most *n* headers — the
   already-taken change 0573 baseline — and strictly fewer for any selective read.

An early-termination refinement is available at zero I/O: a prefix maximum of
`max_end` over the offset-sorted `layout`, one `u64` per record (1,056 bytes for
the 132-record fixture), computed in the central pass that already computes
`all_local_spans_bounded`. Scanning backwards stops at the first index whose
prefix maximum is at or below `offset_t`. The set of headers read is identical;
only the CPU scan shortens.

## Candidates

*(e)* is included because the program's own windowing technique (changes 0565,
0566, 0568, 0570) is the obvious way to keep the full proof and pay less for it,
and it has to be priced before being set aside.

| | Candidate | inv. (1) target | inv. (2) | inv. (3) | inv. (4) | order-independent | converges to (2) |
| --- | --- | --- | --- | --- | --- | --- | --- |
| a | zero-I/O central pre-screen, target-only validation | full | **refutation only** (proves nothing on real packages) | full, free | full, free | yes | no |
| b | target + residual-window neighbours, memoised | full | **target disjointness**, complete per read | full, free | full, free | **yes** | **yes** |
| c | incremental proof, non-overlap only against already-memoised records | full | grows with read history | full, free | full, free | **no** | yes |
| d | full archive-wide proof, one read per header (change 0573) | full | full | full | full | yes | n/a |
| e | full archive-wide proof, coalesced header runs | full | full | full | full | yes | n/a |

Reads and bytes, from `results/change-0575/zip_proof_cost_model.py`. The
archive-wide rows are paid once, in full, on the first read of any member; the
scenario rows are the total for the whole scenario with memoisation.

### `ConditionalFormattingSamples.xlsx` — 132 records, 654,688 bytes

| Candidate | one member (`xl/worksheets/sheet1.xml`) | one-cell closure (7 members) | all 132 members |
| --- | ---: | ---: | ---: |
| today (2 reads/header) | 264 reads / 10,740 B | 264 / 10,740 B | 264 / 10,740 B |
| d — change 0573 | 132 / 10,740 B | 132 / 10,740 B | 132 / 10,740 B |
| e — coalesced, gap ≤ 512 B | 61 / 28,309 B | 61 / 28,309 B | 61 / 28,309 B |
| e — coalesced, gap ≤ 4 KiB | 22 / 91,707 B | 22 / 91,707 B | 22 / 91,707 B |
| e — coalesced, gap ≤ 64 KiB | 1 / 644,889 B | 1 / 644,889 B | 1 / 644,889 B |
| a, c | **1 / 54 B** | **7 / 1,642 B** | 132 / 10,740 B |
| **b** | **4 / 144 B** | **14 / 1,852 B** | 132 / 10,740 B |

### `sheet-names.xlsx` — 13 records, 9,425 bytes

| Candidate | one member (`xl/worksheets/sheet1.xml`) | one-cell closure (6 members present) | all 13 |
| --- | ---: | ---: | ---: |
| today | 26 / 2,478 B | 26 / 2,478 B | 26 / 2,478 B |
| d — change 0573 | 13 / 2,478 B | 13 / 2,478 B | 13 / 2,478 B |
| e — coalesced, gap ≤ 4 KiB | 1 / 8,126 B | 1 / 8,126 B | 1 / 8,126 B |
| a, c | **1 / 54 B** | **6 / 1,592 B** | 13 / 2,478 B |
| **b** | **8 / 264 B** | **11 / 1,742 B** | 13 / 2,478 B |

### `shapes.pptx` — 48 records, 68,822 bytes

| Candidate | one member (`ppt/slides/slide1.xml`) | one-slide closure (6 members) | all 48 |
| --- | ---: | ---: | ---: |
| today | 96 / 4,748 B | 96 / 4,748 B | 96 / 4,748 B |
| d — change 0573 | 48 / 4,748 B | 48 / 4,748 B | 48 / 4,748 B |
| e — coalesced, gap ≤ 4 KiB | 4 / 31,618 B | 4 / 31,618 B | 4 / 31,618 B |
| a, c | **1 / 51 B** | **6 / 1,618 B** | 48 / 4,748 B |
| **b** | **15 / 471 B** | **15 / 1,888 B** | 48 / 4,748 B |

Two corrections to the fixture set. `shapes.pptx` was named as "21 structural
members"; it is **48 central records** (0 directory records, 2 Store, 0 data
descriptors), verified with `zipfile` and again by walking the central directory
with `struct`. The 13 and 132 figures are correct. And per change 0567, "`litchi_xlsx`'s
streaming cell read reaches `IndexedArchive::with_verified_entry_reader`, which
builds a strict layout proof that DOCX and PPTX never trigger" — so the
`shapes.pptx` rows are a `soapberry-zip`-level measurement of reading a member
through `read_entry_to`, not a figure any PPTX facade scenario pays today. They
are retained because the PPTX geometry is the adverse case for (b) and should not
be hidden, not because a facade scenario would improve.

### Reading the table

- (b)'s advantage over today is governed by the archive's byte stride, not its
  member count. On the 654 KB fixture the embedded PNGs push distant records out
  of the 65,559-byte window and one member costs 4 reads instead of 264 (66×). On
  the 68 KB PPTX the whole archive fits inside the window, so one member costs 15
  instead of 96 (6.4×), and on the 9 KB XLSX, 8 instead of 26 (3.3×). This is
  worth stating plainly: **(b) is weakest exactly where the absolute cost is
  already smallest.**
- (b) costs 2× (a)/(c) on a realistic closure (14 vs 7 reads; 15 vs 6) and buys a
  complete, order-independent disjointness proof for every byte returned.
- (e) preserves everything and is a genuine option, but it converts reads into
  bytes at a poor rate on a range source — 22 reads for 91,707 bytes, or 1 read
  for 644,889, to deliver a member whose payload is 1,251 bytes — and it never
  becomes proportional to accessed content. It is the fallback if (b) is
  rejected, not the recommendation.

### Retention and concurrency

| | retained | concurrency |
| --- | --- | --- |
| today | `Vec<StrictEntryLayout>` + `HashMap<u64, usize>` for all *n* records, built on first read | one `Mutex` + `Condvar` single-flight over the **whole** proof: a thread reading member X blocks a thread reading member Y for all 264 reads |
| b | per-record memo of touched records only, plus an optional `u64` prefix-max array (1,056 B at n=132) | per-record memo; disjoint members proceed independently, and the residual sets of two OOXML parts typically do not intersect |
| c | same as b | same as b |
| e | same as today | same as today |

Under (b) and (c) the existing `StrictLayoutCacheState::Building { owner }`
re-entrancy error, the `is_ready` and `build_count` test hooks and the
whole-proof condvar change shape. The failure rule must survive verbatim:
today a failed build resets the cell to `Empty` so a later read retries, and
`GOAL.md` forbids caching an error in a way that blocks recovery after
cancellation, a budget change or a new source version. A per-record memo must
memoise successes only.

## The semantic difference, in bytes

`results/change-0575/overlap_witness.py` builds the archive below and prints the
verdict matrix. It is 4,405 bytes and every member validates in isolation: local
and central names, methods, flags, CRCs and sizes all agree for all three.

```
member    local_off   csize  local_extra  exact_end  cd_min_end  cd_max_end
A.bin             0      32         4096       4163          67       65626
B.bin          2048      32            0       2115        2115       67674
C.bin          4163      32            0       4230        4230       69789
central directory at 4230
```

A's local extra field is a 4,096-byte unknown extra record (header id `0xFACE`,
which every conforming reader ignores), so A's declared local span is `[0, 4163)`
and it **physically contains B's entire local record** at offset 2048. A's
*central* record declares an extra-field length of 0, so the central directory's
bracket for A is `[0, 67)` at minimum and `[0, 65626)` at maximum. The overlap is
invisible to any central-directory-only analysis; only the two bytes at
`offset 28..30` of A's local header reveal it.

| target | today, (d), (e) | (a) | **(b)** | (c) with an empty memo |
| --- | --- | --- | --- | --- |
| `A.bin` | REFUSE | accept | **REFUSE** | accept |
| `B.bin` | REFUSE | accept | **REFUSE** | accept |
| `C.bin` | REFUSE | accept | **accept** | accept |

That last row is the decision the project has to make. **Under (b), an archive in
which two members you never touch overlap each other is read successfully.**
Reading either of the overlapping members still fails, identically to today, from
either side: reading A refuses because B is a successor whose offset lies before
A's exact end; reading B refuses because A is in B's residual window and A's
30-byte header reveals `end_A = 4163 > 2048`.

(c)'s row is worse than it looks. The table shows (c) with an empty memo; the
same call `read("C.bin")` returns `Ok` on a fresh reader and — for an archive
where C overlapped something already read — `Err` on a reader that has read other
members first. Under (a) and (b) each target's verdict is a function of the bytes
alone.

The realistic OOXML shape of this witness: a producer emits
`xl/media/image1.png` with an inflated local extra field so its span swallows
`xl/media/image2.png`. Today, reading `xl/worksheets/sheet1.xml` fails. Under
(b), reading the sheet succeeds and reading either image fails.

### Which existing tests survive

| Test | Construction | Survives (b)? |
| --- | --- | --- |
| `office.rs:12372` `borrowed_store_refuses_an_overlapping_non_target_span` | inflates entry 1's compressed size at local `+18..22` **and central `+20..24`** so its span runs one byte into entry 2's header | **yes**, refuted at zero I/O: the central size makes `min_end_1 = offset_2 + 1` |
| `office.rs:12509` `borrowed_store_refuses_overlapping_local_spans` | rewrites entry 2's central `+42..46` to entry 1's local offset | **yes**, invariant (3), zero I/O |
| `office.rs:11426` `strict_read_to_accepts_prefix_and_rejects_span_intrusion_before_sink` | same size inflation, plus intrusion into the central directory | **yes**, (a)-refutation and per-entry invariant (4) |
| `office.rs:12048` `indexed_strict_layout_accepts_prefix_and_gaps_and_rejects_store_deflate_spans` | 7-byte prefix and 11-byte gap accepted; 2×2 Store/Deflate matrix with the same local+central size inflation | **yes** on all four negative cases; prefixes and gaps stay accepted |

Every overlap test in the repository patches the **central** compressed size,
because it has to — local and central sizes are cross-checked — and is therefore
caught by the zero-I/O refutation, not by reading a local header. **No existing
test constructs the inflated-local-extra-field attack that this record's witness
constructs**, and none asserts on any of the five error strings verbatim. There
is also no test at all for the `"strict streaming refuses duplicate ZIP local
spans"` branch on the strict path; `office.rs:11291` builds equal local-header
offsets but asserts only name ordering and never invokes a strict path.

Any implementation of (b) must therefore add the witness as a fixture and a test
for the duplicate branch, or it will be landing on a test suite that cannot
distinguish it from (a).

## ADR compliance

The house style in `ADR_COMPLIANCE.md` is a prose section per change; a matrix is
used where several alternatives must be compared, as in the change-0497 and
change-0483 tables. This is that case.

| Candidate | 0001 priorities and API layers | 0005 I/O, memory, caching, evidence | 0006 preservation, validation, security | 0010 / 0011 archive and physical-package ownership | Verdict |
| --- | --- | --- | --- | --- | --- |
| **a** zero-I/O screen, target only | Ambiguity remains a typed `Result`; no API surface moves. | Complies: strictly less I/O, fallible allocation unchanged, cache still semantically invisible. | **Weak.** Measured above: proves *no* non-overlap on any real package, so ADR 0006's "ambiguous … ownership fail before publication" posture is reduced to a refutation-only screen for the sake of one read per member. | `soapberry-zip` is the ZIP grammar owner below `litchi-opc`; nothing crosses a crate edge. | **Not recommended.** No ADR conflict, but it discards a real property for a marginal gain over (b). |
| **b** target + residual neighbours | Same. The refusal stays typed and stays before any sink write. | Complies. Reads fall on every selective scenario; the memo records successes only, so a cancelled or budget-failed read still retries, as `GOAL.md` requires. Verdicts are a function of the bytes alone, so "Cache behavior is semantically invisible" holds exactly. | Every byte returned is proven not to lie in another record's declared span, and invariants (3) and (4) are kept in full at zero cost. The delta is confined to records the caller never reads. Preservation, determinism of output, and every limit are untouched. | Unchanged; both readers stay in `soapberry-zip`. | **Recommended, gated.** No ADR exception requested. See the gate below. |
| **c** incremental, memo-relative | Same. | **Conflicts with ADR 0005.** "Semantic payloads load lazily into thread-safe weighted caches… **Cache behavior is semantically invisible.**" Under (c) the same call on the same bytes returns `Ok` or `Err` depending on what the reader read earlier: the memo's contents are the deciding input. | Also weakens 0006's fail-before-publication posture in an order-dependent way, which is worse than (b)'s delta because it cannot be stated as a property of the archive. | Unchanged. | **Conflicts with an accepted ADR. Per `docs/GOAL.md` the conflict is recorded and not implemented.** No proposed ADR is drafted, because (b) obtains the same read-count class without the conflict. |
| **d** full proof, one read per header | Same. | Complies; this is the already-taken baseline. | Unchanged. | Unchanged. | Baseline, not a rival. |
| **e** full proof, coalesced runs | Same. | **Tension, not conflict.** ADR 0005's bounded-resource rule is satisfiable (every run is clamped and fallibly allocated), but the trade is 1 read for 644,889 bytes on a 654 KB archive to deliver a 1,251-byte member, which moves *away* from `GOAL.md`'s proportionality clause while satisfying its syscall clause. Intra-run error precedence changes, the same accepted change as 0565/0566/0568. | Unchanged — every invariant survives byte for byte. | Unchanged. | **Fallback only.** Recommended if and only if (b)'s semantic delta is judged unacceptable. |

Two cross-cutting ADR notes:

- **ADR 0005, "Opening performs container, relationship/catalog, security, and
  mandatory structural validation."** Today's proof is built on *first read*, not
  at open, and is skipped entirely if the caller only opens and lists. Whatever
  the archive-wide proof is, it is already not an open-time structural check, so
  no candidate here moves it out of a position ADR 0005 put it in.
- **ADR 0010's unmeasured-cost clause** — "Representative mixed-format benchmarks
  and profiles must determine whether repeated probing is material… must not
  return merely to reduce an unmeasured cost" — is written about the facade
  dependency edge, but its posture governs this change. It is the reason the
  recommendation below is gated on a measurement rather than stated outright.

## Recommendation

**Implement candidate (b), on both `IndexedArchive<R>` and `ArchiveReader<'data>`
so there is exactly one acceptance contract, gated on first measuring the
range-source scenario. Do not implement (c): it conflicts with ADR 0005.**

Applying (b) to the slice-backed reader as well is a deliberate choice and it has
a cost. That reader pays no I/O, so (b) buys it only CPU (`GOAL.md`'s first
optimization rule, eliminate unnecessary work: 132 header parses and name
comparisons become 14 on the one-cell closure), and it *loses* an archive-wide
property it currently has for free. The alternative — relax the source-backed
reader and leave the slice-backed reader strict — produces a library where the
same bytes are accepted or refused depending on whether the caller supplied a
path or a `Vec<u8>`. One contract with one wording is worth more than a free
property on one ingress; but this is the one sub-decision in this record that a
reviewer could reasonably reverse, and reversing it changes nothing else in the
design.

The gate, and the honest reason for it. Change 0567 already priced this — 4.45 µs
on the 13-member fixture, 42.94 µs on the 132-member one — and declined, and that
judgement is correct for a warm local file: 264 reads of resident page-cache
pages is tens of microseconds. The case that changes the calculus is the one
0567 itself named ("justify such work against a latency-bearing source, not a
warm-cache count") and the one `GOAL.md`'s workstream A names — "Avoid request
amplification from many tiny remote reads" — where 264 requests against a
caller-supplied range source at 1-50 ms per request is 0.26-13 seconds to deliver
one cell, and (b) reduces it to 4 requests. **If that scenario is not measured first, this change
is reducing an unmeasured cost and should not land.**

### Falsification criteria

This recommendation is withdrawn if any of the following turns out to be true.

1. **A rationale exists that this search missed.** If a threat model, a review
   comment, an issue or an out-of-repository record shows the archive-wide scope
   was a deliberate defence against a named attack, the calculus changes
   completely and (e) becomes the recommendation. The archaeology above is a
   negative finding over the repository, not a proof of absence outside it.
2. **The range-source amplification is not material.** If an instrumented
   range-source benchmark over the three fixtures shows the proof is under ~5% of
   the one-cell scenario at realistic latency, this is 0567's 4.45-42.94 µs warm
   saving dressed up, and change 0567's judgement stands.
3. **Target disjointness is not in fact the property that protects the read.**
   The argument above is that the bytes returned for member X are fixed by X's
   own validated local header, so another record's span only matters where it
   intersects X's. If a reviewer produces a concrete case where a non-intersecting
   overlap elsewhere changes what X's read returns or admits, (b)'s premise is
   wrong.
4. **The materializing-read asymmetry is a bug, not a contract.** This record
   leans on `read`/`read_entry` already bypassing the proof. If the project
   decides that is a defect and the proof belongs on every read path, then the
   invariant's reach is about to *grow*, and the right change is to extend (b) —
   which is affordable there — rather than to relax the sink path.
5. **A residual set proves unbounded in practice.** The model bounds residual
   work by *n* and measures the scenario-level screen at 3-14 extra headers, but
   the worst *single* target already reaches 47 of 47 predecessors on
   `shapes.pptx` and 39 of 131 on `ConditionalFormattingSamples.xlsx`. A producer
   corpus whose mean per-target residual approaches *n* would make (b) equal to
   (d) in cost while still weaker in property, at which point (d) or (e) wins.

## Frozen prediction

Stated before anyone measures, so the implementing change can be scored against
it. All figures are positional `read_exact_at` calls issued by the strict-layout
path only, on the source-backed `IndexedArchive<R: ReaderAt>` reader, excluding
payload reads, the EOCD locate and the central-directory scan. Derived from the
fixture geometry by `zip_proof_cost_model.py`; the "today" column is a model, not
a measurement, and the first admission gate is to confirm it by counting.

| Fixture | Scenario | today | with change 0573 | **predicted with (b)** |
| --- | ---: | ---: | ---: | ---: |
| `sheet-names.xlsx` (13) | open, list | 0 | 0 | **0** |
| | read one member | 26 | 13 | **8** |
| | one-cell closure (6 parts) | 26 | 13 | **11** |
| | read all 13 members | 26 | 13 | **13** |
| `ConditionalFormattingSamples.xlsx` (132) | open, list | 0 | 0 | **0** |
| | read one member | 264 | 132 | **4** |
| | one-cell closure (7 parts) | 264 | 132 | **14** |
| | read all 132 members | 264 | 132 | **132** |
| `shapes.pptx` (48) | open, list | 0 | 0 | **0** |
| | read one member | 96 | 48 | **15** |
| | one-slide closure (6 parts) | 96 | 48 | **15** |
| | read all 48 members | 96 | 48 | **48** |

Bytes moved by those reads, predicted with (b): 264 / 1,742 / 2,478 for
`sheet-names.xlsx`; 144 / 1,852 / 10,740 for `ConditionalFormattingSamples.xlsx`;
471 / 1,888 / 4,748 for `shapes.pptx`, against 2,478 / 10,740 / 4,748 for the
whole-archive proof in every scenario. Unlike change 0566's windowing, **this
design reduces bytes as well as reads in every modelled case**; it has no
byte-for-read trade to disclose.

Four predictions that are falsifiable independently of the counts:

- Open and list stay at **zero** strict-layout reads on all three fixtures. That
  is the control proving the delta is the proof and not the index.
- `read`/`read_entry` (the materializing family) stay at **zero** on all three,
  before and after, because they do not reach the proof.
- Reading every member converges to exactly the change-0573 count — no record's
  header is read twice.
- The witness archive in `results/change-0575/overlap_witness.py` is refused for
  `A.bin` and `B.bin` and accepted for `C.bin`, and all three overlap tests named
  above keep their current verdicts and their `ErrorKind::InvalidInput`.

## Admission gates

- **Range-source measurement first.** An instrumented caller-supplied `ReadAt`
  with configurable per-request latency, counting requests and bytes for the
  one-cell scenario on all three fixtures, before and after. Without it, ADR
  0010's unmeasured-cost posture blocks the change.
- **Counted reads.** Tests asserting the frozen table above per scenario, each
  shown to fail without the change, plus the zero-read control for open, list
  and the materializing read family.
- **The witness as a fixture.** `overlap_witness.py`'s archive committed as a
  test fixture with the three-row verdict matrix asserted, since no existing test
  constructs an inflated-local-extra-field overlap.
- **The untested branch.** A test for `"strict streaming refuses duplicate ZIP
  local spans"` on the strict path, which has none today.
- **Order independence.** A test reading the same member from two readers that
  have read different prior members, asserting identical verdicts — the property
  that separates (b) from (c) and the one ADR 0005 requires.
- **Concurrency.** The existing `indexed_strict_first_read_uses_one_single_flight_builder`
  and `indexed_strict_streams_reuse_the_bounded_layout_path` contracts restated
  for a per-record memo, plus a test that two threads reading disjoint members no
  longer serialise.
- **Failure retry.** A test that a read failing for cancellation or a budget
  error leaves nothing memoised and a later read retries, matching today's
  reset-to-`Empty` behaviour.
- **Correctness.** The full `soapberry-zip` and `litchi-opc` suites, the
  `parse_zip` fuzz target, Clippy and rustdoc with warnings denied, a
  no-default-features check, and the OOXML facade suites before commit.

## Retained evidence

Everything in `results/change-0575/` is pure Python over the fixture bytes; no
`cargo` build and no crate under test is involved, so it is reproducible without
the workspace `target/` directory.

| File | What it produces |
| --- | --- |
| `zip_layout_census.py` | Walks each fixture's central directory and every local header with `struct`; reports record counts, actual local extra-field lengths, inter-record gaps, the zero-I/O adjacency verdicts and the per-target residual counts. |
| `layout-census.txt`, `layout-census.json` | Its output for the three fixtures. |
| `zip_proof_cost_model.py` | Models positional reads and bytes for candidates (a)-(e) per scenario, including the 30-byte-neighbour-screen refinement. |
| `proof-cost-model.txt`, `proof-cost-model.json` | Its output; the source of every figure in the candidate tables and the frozen prediction. |
| `overlap_witness.py` | Builds the 4,405-byte inflated-local-extra-field witness archive and prints the per-candidate verdict matrix. `--out` writes the archive itself. |
| `overlap-witness.txt` | Its output. |

Reproduce with, from the repository root:

```
python3 docs/performance/results/change-0575/zip_layout_census.py \
  test-data/ooxml/xlsx/sheet-names.xlsx \
  test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx \
  test-data/ooxml/pptx/shapes.pptx
python3 docs/performance/results/change-0575/zip_proof_cost_model.py
python3 docs/performance/results/change-0575/overlap_witness.py
```

## Limitations

Nothing here is measured. Every read and byte count is a model derived from the
fixtures' own central directories and local headers by the retained Python
scripts; none is a counted syscall, and the "today" column has not been confirmed
against `strace`. No latency, allocation, RSS, cold-cache, concurrency,
cross-platform or producer-corpus result is claimed, and no claim is made that
(b) improves any end-to-end scenario — only that it issues fewer reads in the
model.

The archaeology is a negative finding over this repository at this revision.
`git log -S` finds the commits that introduced the symbols, not discussions that
never reached the repository.

The three fixtures are all from a narrow producer family: all three have
contiguous members with zero gaps, local extra-field lengths drawn from
{0, 40, 264, 520}, no data descriptors and no ZIP64. The residual-window
measurements are therefore characteristic of that family, not of ZIP in general.
An archive with data descriptors or a producer that pads local extra fields would
shift every residual figure, and none of `docs/performance/`'s producer corpora
has been checked for this.

The lower-bound argument assumes an attacker-chosen local extra-field length is
the only unknown. It is sound for the strict path, where local and central names
and sizes are cross-checked, and would not hold for the preservation path
(`validate_preservation_entry_layout`), which deliberately permits a local/central
name mismatch. This design does not touch that path and makes no claim about it.

The design is sequenced after change 0573. It reuses that change's single-read
header, and its per-record memo replaces the whole-proof single flight rather
than composing with it.
