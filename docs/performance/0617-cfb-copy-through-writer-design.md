# 0617: the OLE2 container rebuild is 1.1% of a length-changing XLS save and 13.3-30.0% of a DOC one; the copy-through writer is designed, priced and blocked on an ADR

Status: retained, design only. `performance_claim: none` — the counts and
paired medians below are reported as evidence, not registered as claims. **No
file under `crates/` was modified by this change.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was changed

Nothing in production. This record carries four things:

1. the **first baseline of a length-changing OLE2 edit-and-save**, which change
   0587 recorded as having no selector and size `unknown` (CFB-2, rank 32): a
   retained scratch probe that performs one such save on XLS, DOC and PPT
   through the public editors and attributes it;
2. the **answer to 0587's falsification condition for CFB-2**, which is
   format-dependent rather than yes or no, and therefore splits the item;
3. the **frozen design** of a copy-through writer over a validated
   `SharedOleFile` — its composition shape, its planner, its nine invariants,
   its admission gates and its typed refusals — so that a later batch has it
   without re-deriving it;
4. the **open physical-layout policy question** stated precisely enough for an
   ADR to answer, with the three candidate policies and what each preserves.

Two adjacent findings that are not the subject of this record are reported at
the end: a determinism defect in `OleWriter::write_to`, and the fact that no
`.ppt` fixture in this repository can reach a length-changing shape-text edit.

## The mechanism, restated from the source

`OleWriter` (`crates/litchi-cfb/src/writer/core.rs:250-1222`) is a from-scratch
builder. It holds `streams: Vec<(Vec<String>, Vec<u8>)>` — every stream's
complete payload — and `write_to` re-serializes header, FAT, DIFAT, MiniFAT,
directory and every stream from scratch, allocating every sector afresh in
insertion order through `FatBuilder::allocate_chain`. There is no notion of an
existing physical layout to preserve; `DirectoryBuilder` re-derives the whole
red-black directory tree, including sibling links and node colours, by Apache
POI's `PropertyComparator` order (`writer/directory.rs:530-700`).

Every length-changing OLE2 edit-save in `litchi-xls`, `litchi-doc` and
`litchi-ppt` reaches that builder through one shared path, the generic OLE2
object editor in `litchi-ole-common`:

- `Editor::open` (`object/editor.rs:93-103`) clones the source bytes and parses
  them, then `Package::capture` (`object/codec.rs:23-40`) reads **every** stream
  out through `OleFile::open_stream` — a whole-stream zero-fill plus copy each —
  and retains them all as `Arc<[u8]>`;
- `put_stream_shared` / `put_streams_shared` (`editor.rs:307`, `:364`) clone the
  package, replace the selected streams, and call `commit_candidate_with_rendered`
  (`editor.rs:619-630`), which is the whole container rebuild:
  `package.check` → `render()` (`codec.rs:444-470`, one `OleWriter::create_stream`
  copy per stream, then `write_to`) → `OleFile::open` of the rendered candidate →
  `Package::capture` of the candidate, decoding **every** stream a second time →
  `reuse_stream_allocations`.

So one length-changing save touches every stream's bytes four times — capture,
render copy, emit, recapture — and retains them all at once. The three format
crates reach it from `litchi-xls/src/cell_values/mod.rs:2650`
(`put_stream_shared` of the rewritten `Workbook`),
`litchi-doc/src/tracked_revision/package.rs:2665` (`put_streams_shared` of
`WordDocument` and the selected table stream, which is change 0017's batch), and
`litchi-ppt/src/text_edit.rs:1837-1878` and `slide_order.rs`, both through
`embedded::object::Editor::finish`.

The same-length paths do none of this: `ValidatedOverlayPlan`
(`overlay.rs:569-1006`) streams the source through a 64 KiB window applying
physical spans, never constructs a replacement artifact, and is length-changing
by construction impossible — a `PhysicalSpan` replaces bytes in place inside the
source's own length. Records 0103, 0142, 0143, 0172 and 0175 are all about that
path; 0003 and 0036 are about the builder; nothing covers length-changing
copy-through.

## The baseline

### What was measured, and on what

No registered `perf-baseline` selector performs an OLE2 length-changing save, so
one was built: a scratch Cargo project with path dependencies, retained in the
packet (`results/change-0617/probe/`). It exposes four operations per case so
that callgrind and `perf stat` isolation pairs can difference them:

| operation | what it runs |
| --- | --- |
| `open` | construct the editable snapshot and stop |
| `commit` | `open`, one length-changing edit, and the save — the whole operation |
| `container` | the container rebuild alone: `Editor::open` over the same source plus one `put_streams_shared` of exactly the streams the format crate changed, with exactly the bytes it produced, captured outside the interval. No format record is decoded or re-encoded. |
| `container-changed-only` | the same rebuild over a source holding **only** the changed streams. The difference from `container` is all the container work proportional to the untouched streams — the ceiling of what copy-through can remove. |

Five cases, each a genuine length-changing save through a public editor:

| case | fixture | bytes | edit |
| --- | --- | ---: | --- |
| `xls54016` | `poi/…/spreadsheet/54016.xls` | 984,576 | `cell_values` set A1 of sheet 0 to a string absent from the SST |
| `docfloat` | `ole/doc/FloatingPictures.doc` | 335,360 | `body_text` replace paragraph 0 |
| `docnohf` | `ole/doc/NoHeadFoot.doc` | 26,112 | `body_text` replace paragraph 0 |
| `ppt45543` | `poi/…/slideshow/45543.ppt` | 385,024 | `slide_order` remove slide 1 |
| `pptauthored` | authored by `litchi_ppt::writer::Writer`, 3 slides × 2 textboxes | 8,704 | `text_edit` replace shape text, slide 1 shape 0 |

`pptauthored` is a **generated** fixture and is labelled as such everywhere. It
exists because **no `.ppt` file in this repository can reach a length-changing
shape-text edit**: a sweep of 40 fixtures × 6 targets returns
`Refused(DependencyClosure)` on all 43 resolvable text targets, because real
PowerPoint textboxes carry `ClientTextbox` children that set `can_resize =
false` (`text_edit.rs:1-9`, `:1820-1827`). The real-fixture PPT case is
therefore a structural edit — a slide removal — which is a length-changing OLE2
save on a real producer file.

### Deterministic counts (measured, exact)

`results/change-0617/counts.txt`. Every figure is read back through the ordinary
public CFB parser from the actual output of the actual editor.

| case | source bytes | output bytes | streams | stream bytes in | stream bytes out | unchanged stream bytes | unchanged share |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `xls54016` | 984,576 | 984,576 | 4 | 973,636 | 973,688 | 46,458 | **4.8%** |
| `docfloat` | 335,360 | 342,528 | 13 | 324,540 | 334,948 | 264,190 | **81.4%** |
| `docnohf` | 26,112 | 27,648 | 6 | 23,150 | 24,313 | 12,409 | **53.6%** |
| `ppt45543` | 385,024 | 390,144 | 5 | 379,655 | 384,906 | 64,035 | **16.9%** |
| `pptauthored` | 8,704 | 9,216 | 4 | 5,535 | 6,265 | 128 | **2.3%** |

The changed streams, exactly:

| case | changed streams | before → after |
| --- | --- | --- |
| `xls54016` | `Workbook` | 927,178 → 927,230 |
| `docfloat` | `WordDocument`, `1Table` | 38,775 → 44,032; 21,575 → 26,726 |
| `docnohf` | `WordDocument`, `1Table` | 4,096 → 5,120; 6,645 → 6,784 |
| `ppt45543` | `PowerPoint Document`, `Current User` | 311,524 → 316,775; 4,096 → 4,096 (same length, different bytes) |
| `pptauthored` | `PowerPoint Document`, `Current User` | 5,375 → 6,105; 32 → 32 |

**This table is the single most important result in the record.** The whole
premise of a copy-through writer is that most of an artifact is untouched by an
edit. On a DOC that is true — 81.4% of `FloatingPictures.doc`'s stream bytes are
untouched by a paragraph edit: its `Data` stream alone is 255,308 bytes, and the
seven streams under its `ObjectPool/_1207591883` storage are another 7,449, none
of which a paragraph replacement rewrites. On an XLS it is false:
95.2% of the artifact is the one `Workbook` stream, and every edit rewrites it,
so a copy-through writer has 46,458 bytes to copy through out of 973,636. PPT
sits between: `Pictures` and the property sets are untouched, the `PowerPoint
Document` stream is not.

### Instruction attribution (measured; callgrind, an upper bound)

Isolation pairs at 2 and 12 iterations, per-symbol self costs differenced and
divided by 10, `taskset -c 11`, release build with `debug = 1`
(`results/change-0617/attrib/`, `results/change-0617/attribution.txt`).

| case | whole `commit` Ir/op | `container` Ir/op | container share | untouched-stream work Ir/op | share of `commit` |
| --- | ---: | ---: | ---: | ---: | ---: |
| `xls54016` | 392,345,905 | 16,525,225 | **4.21%** | 11,508,704 | **2.93%** |
| `docfloat` | 23,105,764 | 6,358,582 | **27.52%** | 5,935,753 | **25.69%** |
| `docnohf` | 1,531,433 | 545,924 | **35.65%** | 462,950 | **30.23%** |
| `ppt45543` | 60,090,521 | 5,352,589 | **8.91%** | 3,590,284 | **5.97%** |
| `pptauthored` | 1,492,036 | 312,143 | **20.92%** | 239,757 | **16.07%** |

Ownership of the whole `commit`, by crate, as a share of self instructions:

| case | `litchi-cfb` writer | `litchi-cfb` parser | `ole-common` container | the format crate | runtime (memcpy, memset, SHA, malloc) |
| --- | ---: | ---: | ---: | ---: | ---: |
| `xls54016` | 0.03% | 0.29% | 0.00% | 64.82% | 34.41% |
| `docfloat` | 0.50% | 3.59% | 0.22% | 19.82% | 75.79% |
| `docnohf` | 2.27% | 14.10% | 1.49% | 12.80% | 69.26% |
| `ppt45543` | 0.05% | 2.23% | 0.00% | 1.54% | 96.17% |
| `pptauthored` | 0.82% | 21.14% | 0.00% | 19.59% | 56.23% |

The `litchi-cfb` writer column is the serializer's own arithmetic — FAT,
directory, header construction. It is never above 2.3%. The container *cost* is
not there: it is in the runtime column, and the `container` leg's own profile
says what it is — `__memcpy_avx_unaligned_erms` 63.38% and
`__memset_avx2_unaligned_erms` 21.18% of the `docfloat` container leg. **85% of
the container rebuild is bulk copy**, which is exactly what a copy-through
writer removes and exactly what callgrind over-prices.

The `ppt45543` row is dominated by `sha2::sha256::soft::unroll::compress` at
67.18% — valgrind masks the SHA CPUID bit, so `sha2` runs its software backend.
That figure is roughly five times its native cost and is the reason the PPT rows
are re-priced natively below. It is the PPT persist-directory and publication
hashing, not the container.

### Native cycles (measured)

Callgrind counts `rep movsb` and `rep stosb` once per byte; change 0604 measured
a 35× overstatement for `memset` on this host. The container leg is 85% of
exactly those two instructions, so every callgrind share above is an upper
bound. Priced natively with `perf stat`, isolation pairs at 10 and 110
iterations, median of 11 repetitions each, `taskset -c 11`
(`results/change-0617/cycles.txt`, raw in `results/change-0617/cycles-raw/`):

| case | `open` cycles | `commit` cycles | `container` cycles | `container-changed-only` cycles | container ÷ commit | untouched-stream work ÷ commit |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xls54016` | 60,867,372 | 134,973,045 | 1,444,903 | 497,040 | **1.07%** | **0.70%** |
| `docfloat` | 1,862,311 | 4,956,372 | 656,574 | 27,512 | **13.25%** | **12.69%** |
| `docnohf` | 96,328 | 359,673 | 107,855 | 14,184 | **29.99%** | **26.04%** |
| `ppt45543` | 500,765 | 6,501,647 | 472,088 | 169,054 | **7.26%** | **4.66%** |
| `pptauthored` | 46,751 | 383,970 | 75,473 | 11,235 | **19.66%** | **16.73%** |

| case | `commit` instructions | container ÷ commit (Ir) | untouched ÷ commit (Ir) | callgrind's container ÷ commit |
| --- | ---: | ---: | ---: | ---: |
| `xls54016` | 342,696,585 | 0.45% | 0.29% | 4.21% |
| `docfloat` | 11,950,094 | 10.91% | 10.29% | 27.52% |
| `docnohf` | 1,035,758 | 33.33% | 29.52% | 35.65% |
| `ppt45543` | 10,365,173 | 7.35% | 5.26% | 8.91% |
| `pptauthored` | 1,261,555 | 20.79% | 17.23% | 20.92% |

The last column is the size of the instrument's error on this term. Callgrind
reports 60,090,521 instructions for the `ppt45543` commit where the hardware
retires 10,365,173 — 5.8×, because valgrind masks the SHA CPUID bit and `sha2`
runs its software backend — and 23,105,764 against 11,950,094 for `docfloat`,
1.93×, from `rep movsb` counted per byte. On `docnohf`, where the artifact is
small enough that bulk copy is not the leading term, the two agree to within two
points. **The native columns are the ones to rank on.**

The A/A control, measured in the same window on the same metric: the `docfloat`
commit leg run as both legs, 11 repetitions per level per leg, interleaved with
the runs above. Cycles **+3.29%**; instructions **−0.108%**. So the cycles floor
on this shared host during this window is about 3.3%, which puts `xls54016`'s
1.07% inside it and leaves the other four outside it. The instruction metric is
thirty times tighter, which is why the two are reported together and why the
DOC and PPT conclusions rest on both agreeing.

### Retained bytes at peak (measured)

A counting global allocator in an isolated probe binary, one region per
operation, peak live bytes relative to the region's entry
(`results/change-0617/counts-alloc.txt`). The timing and callgrind binary
carries no allocator wrapper.

| case | source bytes | `open` peak | `commit` peak | `container` peak | `container-changed-only` peak | container peak ÷ source |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `xls54016` | 984,576 | 18,771,707 | 25,410,139 | 5,778,408 | 3,737,315 | 5.87× |
| `docfloat` | 335,360 | 2,454,228 | 3,211,155 | 1,860,994 | 264,546 | 5.55× |
| `docnohf` | 26,112 | 164,258 | 222,239 | 136,535 | 47,366 | 5.23× |
| `ppt45543` | 385,024 | 1,599,900 | 2,735,329 | 2,235,995 | 1,292,277 | 5.81× |
| `pptauthored` | 8,704 | 53,269 | 72,401 | 47,043 | 32,986 | 5.40× |

The container rebuild alone peaks at **5.2× to 5.9× the artifact size** on every
case — several simultaneous copies of the stream set plus the rendered output.
That is the `O(file)` retention 0587 modelled, confirmed and quantified, and it
is the one term where the design's advantage does not depend on the format: the
`container-changed-only` column is what remains when the untouched streams are
absent, and on `docfloat` that is 264,546 bytes against 1,860,994, so **85.8% of
the container rebuild's peak is the untouched streams**. On `xls54016`, where
only 4.8% of the stream bytes are untouched, it is still 35.3%, because the
retention counts the rendered output and the recaptured copy as well as the
original.

## 0587's falsification condition, answered

Change 0587 wrote: *"Falsified if a phase attribution of one XLS/DOC/PPT save
shows the crates' own record re-encoding, not the container copy, dominates."*

The answer is not uniform, and the item has to split:

- **XLS: falsified.** The container rebuild is **1.07% of cycles and 0.45% of
  instructions** of a length-changing cell-value save — inside the 3.29% A/A
  floor measured in the same window — and the work proportional to untouched
  streams is 0.70% of cycles. By callgrind's symbol attribution, `litchi-xls`
  and `litchi-biff` together are 64.82% of the operation's instructions, and its
  largest single terms are `BTreeMap<(u32,u32), Cell>::insert` (11.74%),
  `worksheet::add_cell` (10.60%) and `biff::stream::Records::next` (8.48%) —
  a complete eager re-parse of the workbook during the commit's readback. That
  is item **XLS-9** in the 0587 queue ("XLS edit and save readback is a complete
  eager open", rank 31, size then `unknown`), and this record sizes it: the XLS
  commit costs 134,973,045 cycles against its own open's 60,867,372 — 2.22× —
  and 64.8% of its instructions are the format crate re-reading what it just
  wrote. A copy-through writer would not touch any of it.
- **DOC: not falsified.** The container rebuild is **13.25% of cycles**
  (`docfloat`, 335 KB) and **29.99%** (`docnohf`, 26 KB), of which 12.69% and
  26.04% of the whole save is work proportional to streams the edit never
  touched — four and eight times the A/A floor. `litchi-doc`'s own re-encoding
  is 19.82% and 12.80% of instructions. On the larger fixture the two are
  comparable and on the smaller one the container is larger; on neither does the
  format crate dominate it.
- **PPT: not falsified, but the container is not the leading term either.** On
  the real fixture the container rebuild is **7.26% of cycles** and the leading
  term is whole-artifact SHA-256 hashing — 67.18% of callgrind instructions, and
  the reason the `ppt45543` commit retires only 1.59 instructions per cycle
  where its own open retires 2.21. That is item **DOC-1** in the 0587 queue at
  rank 2, not CFB-2. On the generated single-textbox deck the container is
  19.66% of cycles, but the artifact is 8.7 KB and only 2.3% of it is
  untouched, so that row measures fixed container overhead, not copy-through
  opportunity.

So CFB-2's premise holds where the untouched-stream share is high and fails
where it is low, and the untouched-stream share is a property of the **format's
stream topology**, not of the container. The ranking consequence is stated under
"Limitations".

## The frozen design

### Where it sits

`litchi-cfb` has three publication paths today. The copy-through writer is a
fourth, and replaces none of them:

| path | shape | length-changing? | retains |
| --- | --- | --- | --- |
| `OleWriter` (`writer/core.rs`) | from-scratch builder, in-memory model | yes | every stream |
| `SequentialOleWriter` (`writer/sequential.rs`) | forward-only, predeclared layout, streams each payload from a `Read` | yes (creation only) | bounded metadata plus one buffer |
| `ValidatedOverlayPlan` (`overlay.rs`) | physical spans over a validated `SharedOleFile` | **no** | the spans only |
| **copy-through (this design)** | spans plus an appended tail over a validated `SharedOleFile` | yes | the changed streams plus rebuilt metadata |

### Part 1 — the composition shape: spans plus a tail

The overlay already owns the right primitive. `ComposedOverlaySource`
(`overlay.rs:583-628`) is an immutable `ReadAt` view of *source bytes with
sorted physical spans applied*, and `finish_overlay_plan_with_owner`
(`overlay.rs:1022-1079`) opens `SharedOleFile::open(composed)` before any plan
is returned. The design generalizes that view from same-length to
length-changing by adding a sector-aligned tail:

```rust
pub struct ComposedCopyThroughSource {
    source: SourceSnapshot,      // the original artifact; its bytes are never rewritten
    spans:  Arc<[PhysicalSpan]>, // in-place replacements below source.length
    tail:   Arc<[u8]>,           // appended whole sectors, at source.length and beyond
    version: SourceVersion,
}
```

`len()` returns `source.length + tail.len()`. `read_at(offset, out)` serves
`[0, source.length)` exactly as today — `source.read_at` then `apply_spans` —
and `[source.length, len)` from `tail`, splitting a straddling read into the two
halves. With an empty tail it is byte-for-byte today's `ComposedOverlaySource`,
which is the property that lets the two share one implementation and one set of
tests.

Two consequences that matter:

- the source's own bytes are read but never rewritten, so the source fingerprint
  is computed over `[0, source.length)` exactly as today and compared against
  the same retained value;
- the target fingerprint is computed over the whole composed artifact, so it is
  the spans pass plus one hash of the tail. No byte is read twice.

### Part 2 — the planner

```rust
impl SharedOleFile {
    pub fn plan_copy_through(
        &self,
        replacements: Vec<StreamReplacement>,  // { path: Vec<String>, bytes: Arc<[u8]> }
        limits: CopyThroughLimits,
    ) -> Result<ValidatedCopyThroughPlan, OverlayError>;
}
```

Every step runs over the `ParsedOleIndex` the open already validated. The source
is never re-parsed.

1. **Resolve and bound.** `check_source_version`; resolve each path through the
   validated directory with `path_refs` and `find_entry`; refuse a path that is
   not `STGTY_STREAM`, a duplicate SID, or an empty component — the same
   refusals and the same messages `plan_same_length_stream_overlays`
   (`overlay.rs:655-700`) already raises. Bound by `max_streams`,
   `max_replacement_bytes` and `max_output_bytes` before any allocation
   proportional to the artifact.

2. **Classify placement per changed stream.** Target table is MiniFAT when
   `new_len < mini_stream_cutoff` (4,096) and FAT otherwise — **the cutoff
   migration is decided here, once, and nowhere else**. Old chain via
   `collect_chain_exact(table, start, ceil(old_len / unit), name)`, the helper
   `splice.rs` and `stream_move.rs` already use, so a malformed chain refuses
   with the identical message. New chain length is `ceil(new_len / unit)`.

3. **Build the free-sector ledger.** The open already computed
   `sector_roles: Vec<PhysicalSectorRole>` (`file.rs:152-161`, one byte per
   physical sector, filled by `validate_stream_allocations`). The planner starts
   from it: a sector owned by a changed stream that its new chain no longer needs
   is *released*; a sector whose FAT entry is `FREESECT` is *pre-existing free*;
   everything else is *pinned*. Pinned sectors are never written.

4. **Allocate, then rebuild only the metadata.** Whichever allocation policy is
   chosen (Part 4), the planner then produces: a complete new FAT image; a new
   MiniFAT image and mini-stream chain if any mini stream changed; a new
   directory image; a new DIFAT if the FAT sector count crosses 109; and a new
   512-byte header. **The directory image is built by copying each parsed entry's
   64 bytes verbatim and overwriting only `start_sector` and `stream_size` on the
   changed entries** (plus the root entry's mini-stream start and size). Names,
   name lengths, types, node colours, `sid_left`/`sid_right`/`sid_child`, CLSIDs,
   state bits and both timestamps are never re-derived. This is the sharpest
   difference from `OleWriter`, which rebuilds the red-black tree from scratch,
   and it is what makes ADR 0026's SID binding hold trivially: no SID moves.

5. **Emit as spans and a tail.** Each rebuilt sector whose physical index is
   below the source's sector count becomes a `PhysicalSpan` at
   `sector_offset(sector, sector_size)`; each one beyond becomes tail bytes, in
   ascending physical order with no gaps. `validate_and_coalesce_spans`
   (`overlay.rs:1253`) sorts, checks for overlap and merges adjacent spans, as
   it does today.

6. **Reopen before a sink observes a byte.** `SharedOleFile::open(composed)` —
   the 0175 policy, verbatim from `overlay.rs:3-6`: *"Plans are derived from
   stream chains already validated by `SharedOleFile`, then the composed
   artifact is reopened through the normal CFB parser before a sink can observe
   a byte."* Then two readbacks: every **changed** stream must equal its
   replacement, and — the obligation copy-through adds — every **untouched**
   stream must equal its source bytes. The untouched readback is what turns
   "copy-through" from an assumption into a proof. It is bounded by
   `limits.max_verified_bytes`; above that bound the plan is refused, not
   silently unverified.

7. **Fingerprint bracket.** `finish_overlay_plan_with_owner`'s structure is
   unchanged: two complete source/target fingerprints bracket the candidate
   validation; 0175's owned-immutable specialization applies unchanged, because
   it is a property of the source, not of the composition.

### Part 3 — publication

`write_to<W: Write>` and `save<P: AsRef<Path>>` are `overlay.rs:844-1006`
extended by one loop:

```text
for each 64 KiB chunk of [0, source.length):
    read source; hash into source_hasher; apply_spans; hash into target_hasher; write
for each 64 KiB chunk of tail:
    hash into target_hasher; write
compare observed source fingerprint with the retained one
compare observed target fingerprint with the retained one
```

The whole fingerprint ledger frozen by 0103, 0143 and 0175 survives unchanged:
planning brackets validation with two complete fingerprints; direct `write_to`
keeps its initial preflight, its emission hashing and its post-emission
preflight; atomic `save` keeps its initial preflight, its emission hashing and
its post-flush/fsync pre-rename preflight, and skips only the duplicate
post-emission scan (0103's 4N → 3N); an `open_owned` source skips only the two
outer complete scans (0175). The ≤1 MiB fingerprint window and the 64 KiB
publication window still never coexist (0143's rule).

### Part 4 — the nine invariants

1. **Sector partition.** Every physical sector in `[0, new_sector_count)` is
   claimed exactly once — by the header prefix, a `FATSECT`, a `DIFSECT`, the
   directory chain, the MiniFAT chain, the mini-stream chain or exactly one
   stream chain — and every unclaimed sector has FAT entry `FREESECT`. The
   reopen in step 6 proves it, because that is literally what
   `validate_stream_allocations` and `validate_physical_sector_layout`
   (`file.rs:1110-1224`) check. The planner must assert it *before* emitting, so
   that a planner bug is a typed refusal rather than a candidate that fails to
   reopen.
2. **Untouched-stream byte identity.** Every stream outside the replacement set
   reads back byte-identical. `docs/GOAL.md` rule 2; ADR 0006's preservation
   default. Proved by readback, not by construction.
3. **Directory identity.** Only `start_sector` and `stream_size` of changed
   entries may differ from the source's 64 bytes. Red-black order and colour are
   preserved because they are copied, not re-derived. ADR 0026's SID binding
   holds because no SID moves; ADR 0026's *"state bits and timestamps are
   intentionally not fabricated"* holds because neither field is written.
4. **FAT and DIFAT counts and markers.** `num_fat_sectors` equals the count of
   `FATSECT` entries; `num_difat_sectors` equals the count of `DIFSECT` entries;
   the header's first 109 FAT sector ids followed by the DIFAT chain enumerate
   exactly those sectors in order, terminated by `ENDOFCHAIN`.
5. **v3 size masking.** On a 512-byte-sector file a directory entry's
   `stream_size` high word is written as zero, matching what
   `mask_v3_stream_size` (`file.rs:3029`) tolerates on read; the fixture
   `ole/doc/cfb-v3-uninitialized-size-high-word.doc` is the regression case. A
   stream reaching 2 GiB on v3 is refused with `validate_stream_size`'s existing
   message.
6. **Mini-stream cutoff migration.** A stream crossing 4,096 changes table in
   either direction. Growing: its mini-sectors are released, its bytes move to a
   FAT chain, the mini-stream shrinks, the MiniFAT is truncated and the surplus
   mini-sectors become `FREESECT`. Shrinking: the reverse. Either way the
   mini-stream's own chain stays **exact** — the root entry's `stream_size` is
   `mini_sector_count × 64` and its FAT chain is exactly
   `ceil(that / sector_size)` sectors — because `read_chain_into` refuses a
   chain that "ends before its declared length" or "exceeds its declared
   length".
7. **Deterministic output.** Two plans over the same source and the same
   replacements produce byte-identical output. This forbids `HashMap`/`HashSet`
   iteration order anywhere in the allocator, forbids consulting a clock, the
   filesystem or an RNG (ADR 0006), and is not a hypothetical constraint —
   see the defect reported at the end of this record.
8. **Reopen before a sink observes a byte.** Step 6; the 0175 policy.
9. **No caller-selected physical offsets.** `overlay.rs:3`. The public API takes
   logical paths and replacement bytes; every offset is derived from the
   validated index. This is what keeps the writer policy-neutral in the sense
   0100 requires and stops a format owner from inventing a layout.

### Part 5 — the open policy question: `FREESECT` reuse versus append

Three candidate policies:

- **(A) Append-only.** Never reuse a freed sector. New sectors go after the
  source's last sector; released sectors become `FREESECT` and stay. The
  source's physical layout is then a *prefix* of the output's.
- **(B) Reuse-then-append.** Fill `FREESECT` runs in ascending physical order
  first, then append. Output size is bounded by the high-water mark, but a
  sector that held stream X's bytes may now hold stream Y's.
- **(C) Reuse only the sectors this operation itself released, then append.**
  No pre-existing `FREESECT` is touched, so whatever a producer left in the free
  list stays exactly where it is.

What is settled: **no ADR addresses the physical sector layout of a saved CFB.**
The only ADR sentence naming sectors in a preservation context is ADR 0008's *"A
changed CFB preserves sector size, root and exposed storage CLSIDs, stream
bytes, and hierarchy"* — sector *size*, not placement. ADR 0006's `Preserve`
default lists *"Untouched entries, streams, ordering, compression, timestamps,
unknown markup/records, namespace choices, and lexical details"*, where
"ordering" is directory ordering. And ADR 0005 contains the clause that
*authorizes* this whole design — *"Preserve-mode save raw-copies unchanged
compressed ZIP entries or CFB streams when possible"* — while saying nothing
about where the copies land.

What is not settled, and what the ADR must decide:

- all three policies preserve every logical byte and are indistinguishable to
  `OleFile` and `SharedOleFile`; the difference is visible only to a
  byte-comparing differ;
- (A) is the only policy under which "every untouched sector is byte-identical
  at the same offset" is true, which makes the spans-plus-tail composition its
  natural representation rather than an encoding trick;
- (A) is also the only policy that grows the file on every save. For `docfloat`
  the changed streams are 70,758 bytes, so an append-only save adds about 70 KiB
  per edit to a 335 KiB file — 21% growth per save, unbounded over a session.
  That is why the ADR question cannot be deferred past the first
  implementation;
- (B) is the only policy that can make the output *smaller* than today's
  `OleWriter` output on a shrinking edit;
- (C) bounds the growth to the net size change while leaving a producer's own
  free list untouched.

**The design recommends (C)**, with append as the fallback when the released run
is too small, and asks the ADR to state explicitly that a saved CFB's physical
sector placement is not part of the preservation contract — only sector size,
stream bytes, hierarchy, directory metadata and CLSIDs are.

Until that clarification exists, this design is frozen and must not be
implemented: every one of the three policies is a defensible reading of ADR 0006
today, and choosing one in code would be making preservation policy in a
performance change.

### Part 6 — the admission gates

The writer admits an operation only when every gate holds. Otherwise it returns
a typed refusal **before any byte is emitted**, and the caller falls back to
today's `OleWriter` path, which stays exactly as it is:

1. the source opened through `SharedOleFile`, i.e. a positional source with a
   fully validated `ParsedOleIndex`;
2. every replacement path resolves to an existing `STGTY_STREAM`;
3. no storage or stream is created, deleted, moved or renamed — those change the
   directory tree shape and the SID assignment, which is `OleWriter`'s job and
   ADR 0026's binding. `stream_move.rs` already owns directory-only moves and is
   not affected;
4. replacement count and aggregate bytes within `CopyThroughLimits`, and the
   untouched-stream readback within `max_verified_bytes`;
5. the resulting artifact within the v3 2 GiB ceiling and `validate_output_size`;
6. the source declares **no DIFAT sector**. Zero of the 212 OLE2 fixtures in
   this repository declare one (0587's census), so a first implementation that
   rebuilt the DIFAT would ship untested; refusing is honest and the gate can be
   lifted when a fixture exists;
7. the source is not signed, encrypted or DRM-marked —
   `reject_protected_shared_container`, exactly as
   `SourceBackedOverlayPublisher::open` applies it today
   (`litchi-ole-common/src/source_backed_overlay.rs:38-57`).

### Part 7 — the parity gate the implementation must pass

1. a differential over every OLE2 fixture under `test-data/`: for each, perform
   one length-changing replacement of every stream in turn through both the
   `OleWriter` path and the copy-through path, and compare **every** stream's
   bytes, the directory metadata of every entry, and the CLSIDs, between the two
   outputs and against the source;
2. an idempotence check: the same plan emitted twice is byte-identical, in two
   separate processes (invariant 7);
3. a reopen check on every output through `OleFile`, `SharedOleFile` and the
   owning format crate's reader;
4. the malformed corpus compared by `OleError` and `OverlayError` `Display`
   string, so that every refusal that exists today still fires with the same
   message and in the same order;
5. the cutoff-migration cases in both directions, including a stream that
   crosses 4,096 while another shrinks below it in the same plan;
6. the existing overlay, splice and stream-move suites, unchanged, since the
   composed source is shared with them;
7. `cargo fmt`, `clippy`, `cargo test` and `cargo doc` on `litchi-cfb`,
   `litchi-ole-common`, `litchi-doc`, `litchi-ppt` and `litchi-xls`.

### What the design does not touch

`OleWriter` and `SequentialOleWriter`; the same-length overlay, splice and
stream-move paths beyond sharing their helpers; every ADR 0005 mandatory
validation at open; the format crates' own re-encoding and readback, which is
where the XLS cost actually is; the whole-artifact hashing that dominates the
PPT save, which is 0587's DOC-1 at rank 2.

## Why it is sound

**Invariants.** The nine above. Eight of the nine are proved by the reopen in
planning step 6, because `SharedOleFile::open` already performs exactly those
structural checks on every artifact it parses; the ninth, determinism, is a
property of the planner's own data structures and is proved by the idempotence
check.

**Error identity.** Every refusal the design raises is one that exists today,
raised from the same helper: `collect_chain_exact`'s chain-length pair,
`read_chain_into`'s "Sector N is outside the file", `validate_stream_size`'s
v3 ceiling, `path_refs`' path refusals, `find_entry`'s lookup failures, and
`OverlayError::{SourceFingerprintChanged, TargetFingerprintChanged,
IncompleteOutput, Committed, Allocation, Unavailable}`. The design adds exactly
one new refusal kind — the admission gates — and it fires before any byte is
emitted, which is ADR 0005's *"A changed owned source is publishable only
through a proven preservation plan; if physical framing or opaque members cannot
be preserved, publication returns a typed capability refusal before output."*

**ADR reading.** ADR 0005 authorizes the mechanism (*"Preserve-mode save
raw-copies unchanged … CFB streams when possible"*), keeps its mandatory
structural validation (the reopen performs it), keeps its bounded resources
(`CopyThroughLimits`, and the peak drops from O(file) to O(changed streams +
buffer)), and keeps its atomic publication (`save` is unchanged). ADR 0006's
preservation default is strengthened, not weakened: untouched streams become
byte-identical *by proof* rather than by reconstruction. ADR 0026's directory
binding holds because no SID moves and no state bit or timestamp is written.
ADR 0003's revision binding is not engaged: the writer publishes bytes, not
revisions. The one ADR question the design cannot answer for itself is physical
layout, and it is raised as an explicit clarification request rather than
decided in code.

**`docs/GOAL.md` rules.** Rule 2 (preserve is the default) is the design's
central invariant. Rule 10: no new `unsafe` — the composition is slices and
`Arc<[u8]>`. Rule 11: no public leakage of archive types, raw locks or
executors; `ValidatedCopyThroughPlan` and `ComposedCopyThroughSource` expose the
same surface `ValidatedOverlayPlan` and `ComposedOverlaySource` do today.
Rule 12: `try_reserve` before every allocation, every limit typed, no
malformed-input defence weakened. The optimization order is respected: this is
step 1 and 2 work — unnecessary decode, copy and retention — and no layout,
algorithm, parallelism or SIMD change is proposed.

**Which contracts are untouched.** Every public signature in `litchi-cfb`,
`litchi-ole-common`, `litchi-xls`, `litchi-doc` and `litchi-ppt`. No output byte
of any existing path changes, because the copy-through writer is additive and
gated; a refused operation takes exactly today's route and produces exactly
today's bytes.

## Measured

Everything under "The baseline" is this section's content and is not repeated.
In tiers:

- **measured** — the stream inventories, changed-stream byte counts and
  unchanged-stream shares, read back through the public parser from real editor
  output; the callgrind self-instruction attribution of all four operations on
  all five cases; the native `perf stat` cycles and instructions for the same;
  the allocation regions and their peaks; the A/A floor.
- **modelled** — the realizable saving. `container − container-changed-only`
  bounds it from above: it is *all* the container work proportional to untouched
  streams, where a copy-through writer still writes those bytes once to the sink
  and still walks their chains during the candidate reopen. The design removes
  three of the four byte passes (capture decode, render copy, recapture decode)
  and keeps the fourth (emit), so the realizable saving is roughly three
  quarters of the bound. That fraction is arithmetic over the mechanism, not a
  measurement.
- **unknown** — what the writer costs. It was not implemented, so no `after` leg
  exists for any scenario.

Scope of every figure: scenario as named; corpus the five cases named; machine
AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, CPU 11, with seven other
agents building and measuring concurrently; build `--release` with `debug = 1`,
rustc 1.95.0; metrics callgrind self `Ir`, `perf stat` `cycles` and
`instructions`, and counting-allocator peak live bytes.

## Correctness evidence

No production code changed, so there is nothing to test that was not already
tested. What was run, at the base commit on the clean worktree:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-cfb -p litchi-ole-common --all-targets --locked` | clean (workspace lints are `deny`) |
| `cargo test -p litchi-cfb --release --locked` | 322 + 13 + 6 passed, 12 doctests passed and 1 ignored, 0 failed |
| `cargo test -p litchi-ole-common --release --locked` | 129 + 14 + 2 + 8 + 5 passed, 0 failed |
| `cargo doc -p litchi-cfb -p litchi-ole-common --no-deps --locked` | clean (rustdoc lints are `deny`) |

The two crates are the ones the design would touch. The gates were run on the
clean worktree at the base commit; the only tracked files this change adds are
under `docs/`.

The probe reads fixtures and the public editor APIs only; it modifies no tracked
file and writes nothing into the repository. Its source is retained in the
packet.

## Validation preserved

Nothing moved. Every ADR 0005 mandatory validation runs where it ran; every
typed limit, budget and cancellation point is where it was; no refusal moved
earlier or later; no output byte changed anywhere. This record adds a design,
five baselines and two reported findings.

## Limitations

- **Five cases, one host, one build.** Two DOC fixtures, one XLS, one real PPT
  and one generated PPT. The unchanged-stream share — the quantity that decides
  whether the design is worth anything — ranges from 2.3% to 81.4% across those
  five, so five is not enough to characterize any format, let alone the corpus.
  A census of the unchanged-stream share over all 212 OLE2 fixtures is the
  cheapest next measurement and was not taken.
- **The largest DOC that opens is 335 KB.** `body_text::Snapshot::open` refuses
  the 1.6 MB fixture with `Refused(DrawingDependency)` and most others with
  `invalid stylesheet: style names and aliases must be unique`; of the DOC
  fixtures above 40 KB, two open. No claim about DOC at scale follows.
- **The PPT shape-text case is a generated fixture.** No real `.ppt` here can
  reach a length-changing shape-text edit; the real-fixture PPT case is a slide
  removal instead. Both are length-changing OLE2 saves, but neither is "a
  representative PowerPoint text edit on a producer file", because this
  repository contains no such thing.
- **The `container` leg is an approximation of the container's share.** It runs
  `Editor::open` plus one `put_streams_shared` over the same source with the
  same replacement bytes — the same container work a format commit performs —
  but as a separate operation, so it pays its own source parse and its own
  allocator warm-up. It is not a subtraction of two measured phases of one
  operation, which change 0036 explicitly forbids treating as additive
  components.
- **`container-changed-only` is not the copy-through writer.** It is the same
  rebuild over a smaller artifact, so its difference from `container` also
  contains the smaller FAT, directory and header work, not only the untouched
  streams' bytes. It bounds the saving; it does not measure it.
- **Callgrind over-prices this particular term more than most.** 85% of the
  container leg is `rep movsb`/`rep stosb`, counted once per byte; 0604 measured
  a 35× overstatement for `memset` on this host. The native cycles table is the
  figure to rank on, and the callgrind table is retained because it is the only
  one that attributes by symbol.
- **Not claimed:** any speedup, any regression, any peak-RSS, cold-cache,
  file-backed, range-source, syscall, throughput or cross-platform result; any
  statement about a format outside the five cases; any conclusion about which of
  the three layout policies is correct; any size for the copy-through writer
  itself.
- **What is left open.** The ADR clarification on physical layout, which blocks
  implementation. The corpus-wide unchanged-stream census, which decides whether
  CFB-2 is worth a rank at all. And the finding this baseline turned up as a
  side effect: on XLS the container is 1.07% of cycles while the format crate's
  own re-parse during the commit readback is 64.8% of instructions, which is
  0587's XLS-9 at rank 31 and is now the larger, cheaper, lower-risk item of the
  two.

## Adjacent findings, reported and not fixed

**1. `OleWriter::write_to` output is not deterministic when a document has two
or more explicitly created storages.** `write_to` iterates
`self.storages: HashSet<Vec<String>>` and `self.storage_clsids: HashMap<…>` to
call `DirectoryBuilder::add_storage_path` (`writer/core.rs:977-982`), and
`add_storage_path` assigns each new entry `sid = self.entries.len()`
(`writer/directory.rs:392`), with `generate_directory_stream` serializing
entries in SID order (`directory.rs:573-577`). Rust's default hasher is seeded
per process, so the SID assigned to each storage — and therefore the directory
image, and therefore the whole file — depends on hash iteration order across
runs. Streams are unaffected (`self.streams` is a `Vec`), and a document whose
storages are all created implicitly by `add_stream_path` is unaffected, because
that path visits them in stream insertion order. ADR 0006 requires
*"Serialization is deterministic unless a `Clock`, actor identity, or
cryptographic RNG is explicitly supplied."*

Confirmed by execution, not only by reading: a probe that builds a CFB with *N*
explicitly created storages, each holding one stream, and prints an FNV-1a
digest of the bytes, run 12 times in 12 separate processes per *N*
(`results/change-0617/determinism.txt`) — **1 storage: one digest in 12 runs;
2 storages: two distinct digests; 3 storages: six; 8 storages: twelve distinct
digests in twelve runs**, with the output length identical every time. This is a
correctness finding on a path this record did not change; it is reported for the
owner, not fixed here, and invariant 7 of the design exists because of it.

**2. No `.ppt` fixture in this repository can reach a length-changing shape-text
edit.** 40 fixtures × 6 `(slide, shape)` targets: 43 `DependencyClosure`
refusals, 28 `NoTextAtom`, 29 `ShapeNotFound`, 8 `NoTextbox`, 24
`UnsupportedEncryption(CryptoApi)`, 48 `SlideNotFound`, and zero admissions
(`results/change-0617/ppt-target-sweep.txt`). Every `litchi-ppt` test that
exercises the length-changing text path authors its deck with
`litchi_ppt::writer::Writer`. This is the same shape of finding as change 0607's
"no opened deck can reach the PPTX slide regeneration": a landed capability with
no corpus that exercises it. It is reported, not fixed.

## Retained evidence

[`results/change-0617/README.md`](results/change-0617/README.md) — the probe
source, the capture scripts, the deterministic counts, the callgrind
attributions and their raw profiles, the native cycle figures and their raw
`perf stat` output, the allocation regions, the PPT target sweep, the gate
tails, `decision.json` and `log-sections.md`.
