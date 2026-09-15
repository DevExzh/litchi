# 0579: resume the CFB allocation-chain walk instead of restarting it at sector zero

Status: retained. `performance_claim: none` — this record carries paired medians
in two directions across two windows with a measured noise floor, isolated
hardware counters, per-open chain-step counts and a corpus-wide I/O differential,
not a registry claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements opportunity 3 of change
[0574](0574-ole2-next-opportunity-survey.md). It does **not** implement it the
way that record proposed. The mechanism 0574 named — giving the globals scan a
retained `SharedOleStreamCursor` — was built first, measured, and **rejected**
because it made the open *slower*. The record's most useful finding is why.

## What was removed

`SharedOleFile::read_stream_range` reaches a stream offset by walking the
allocation chain from the stream's **first** sector, on every call:

```rust
let mut sector = start_sector;
let first_ordinal = usize::try_from(offset / sector_size as u64)?;
for _ in 0..first_ordinal {
    sector = next_chain_sector(&self.index.fat, sector, "FAT")?;
    ...
}
```

`GlobalsBuffer` (`litchi-xls/src/workbook/source.rs`) fills the workbook globals
forward in a doubling window — 20 calls on the flagship fixture, each starting
where the last one ended — so the prefix walk is quadratic in the number of
fills. Change 0574 counted it: **5,796 `next_chain_sector` calls per open** of
`ConditionalFormattingSamples.xls`, where the bytes require about 1,077 links.

The walk is now resumable. `litchi-cfb` gains one `Copy` value,
`StreamChainHint`, obtained from `SharedOleFile::chain_hint()`, and one method,
`read_stream_range_hinted`, which begins the prefix walk at the position the hint
retains instead of at sector zero and records where it finished. `GlobalsBuffer`
keeps one hint for the whole scan.

**There is one implementation of the read.** `read_stream_range` is now
literally `read_stream_range_hinted` with a fresh hint, so the read loop, the run
batching, every bound check, every refusal and the source-version fence are the
same code on both paths. Nothing was duplicated — the lesson change
[0570](0570-cfb-fat-run-batching.md) recorded when it declined to reuse
`read_sectors_batched` was that a second copy of a validated loop is how error
identity drifts.

The MiniFAT range reader is threaded the same way. The unhinted entry point
passes a hint recorded under a reserved SID that no directory entry can carry, so
it behaves exactly as it did before hints existed.

### The bounded-memory argument

A hint is a borrow of its reader plus a sector ordinal, a sector index, a first
sector, a SID and a flag. It is `Copy`, it allocates nothing, and its size does
not depend on the stream: it is a **position**, not an index of the chain. That
is the same argument change 0570 made for its run scratch buffer, and it is
easier here because there is no buffer at all. Retaining a materialised chain
index was never considered: `ParsedOleIndex` already holds the whole FAT, so an
index of one stream's chain would be a second copy of data the reader already
has.

## Measured effect

Two binaries built from a **detached git worktree of `32d25e088` outside the
repository working copy**, each with its own `CARGO_TARGET_DIR`, differing only
in `crates/litchi-cfb/src/{shared.rs,lib.rs}` and
`crates/litchi-xls/src/workbook/source.rs`. Identities are in
[`environment.json`](results/change-0579/environment.json). The driver is change
0574's `capture_counters.sh` with one line changed — the scratch directory it
exports as `TMPDIR` — so children are launched exactly as that record's were:
CPU 17 pinned, ASLR disabled, single-threaded.

**Host quiescence is not established.** Other agents were active throughout; the
load average moved between 0.7 and 8.0 on 32 cores. That is why the noise floor
below is measured in the same window rather than assumed, and why every latency
figure is reported twice in two windows.

### The control: not one byte of I/O moved

Every logical counter is identical between the legs, in all 18 captured cells,
and constant across all 100 samples of every cell:

| mode / operation | reads | read bytes | `version()` | `len()` |
| --- | ---: | ---: | ---: | ---: |
| any mode / open | 53 | 565,201 | 29 | 1 |
| any mode / list | 53 | 565,201 | 29 | 1 |
| any mode / one-cell | 61 | 603,115 | 42 | 1 |

The open row reproduces change [0565](0565-xls-globals-single-pass.md)'s and
change 0574's retained figures exactly.

Corpus-wide, one source-backed open of **every** `.xls` and `.xlt` fixture under
`test-data` through a range-recording positional source:

| | before | after |
| --- | ---: | ---: |
| fixtures | 126 | 126 |
| opened | 119 | 119 |
| refused | 7 | 7 |
| positional reads | **1,858** | **1,858** |
| bytes read | **4,419,376** | **4,419,376** |

The two logs are **byte-identical**, and they carry more than the totals: per
fixture they record the source-version count, an FNV-1a digest of the **exact
positional read ranges in order**, and, for the seven refused fixtures, the
refusal text. Nothing in that log moved. This change removes CPU work only.

### Chain steps per open

Callgrind isolation pairs — differencing a large-sample and a small-sample child
and dividing by the sample delta, change 0564's method — annotated at
`--threshold=100` so no small edge is dropped.

| fixture | chain links before | after | removed | whole-open Ir before | after | |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | **5,796** | **2,099** | 3,697 | 2,537,912 | 2,380,011 | **−6.22%** |
| `SimpleWithImages-mac.xls` | 429 | 297 | 132 | 341,033 | 335,112 | −1.74% |
| `54016.xls` | 1,995 | 1,110 | 885 | 7,313,414 | 7,239,616 | −1.01% |
| `WithCustomViews.xls` | 456 | 336 | 120 | 922,966 | 915,765 | −0.78% |
| `WithExtendedStyles.xls` (mini) | 32 | 32 | 0 | 124,596 | 123,858 | −0.59% |
| `SimpleWithColours.xls` (mini) | 81 | 73 | 8 | 157,983 | 157,670 | −0.20% |
| `SimpleMultiCell.xls` | 4 | 4 | 0 | 138,367 | 138,435 | +0.05% |
| `WithCheckBoxes.xls` (mini) | 70 | 62 | 8 | 216,535 | 216,840 | +0.14% |

**The flagship's before figure is 5,796, reproducing change 0574's
independently captured 5,796 exactly**, which is the strongest available evidence
that the two records are counting the same thing. The after figure, 2,099, is not
the ~1,077 change 0574 said the bytes require, and should not be. The profile
attributes both of `read_stream_range_hinted`'s loops to one symbol, but the
rejected cursor leg below separates them by construction, because a cursor walks
each link exactly once: it measures **1,076**. The 1,023 links the hint form
still walks are therefore the read loop's own run discovery inside the requested
ranges — that work is the read, not a prefix, and the hint neither touches it nor
should.

**Two fixtures move by less than 0.15% and are reported rather than rounded
away.** `SimpleMultiCell.xls` and `WithCheckBoxes.xls` gain 0 and 8 chain steps,
and their instruction counts rise by 68 and 305 per open. That is the hint's own
bookkeeping — constructing it once and comparing four fields on each of the
fills — on fixtures where there is nothing for it to save. It is below the level
at which an isolation pair is reliable, and neither fixture has a wall-clock
measurement.

**Three of the eight fixtures are mini-stream resident**, which is the case
change 0574 named as this change's falsification test. They behave as expected
rather than differently: the MiniFAT reader takes the same hint, saves 8, 8 and 0
links, and moves no read.

### Instructions and cycles per open

`perf stat`, isolated the same way, on the `open` operation.

| fixture | mode | instructions | | cycles | | IPC |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | owned | 1,389,409 → **1,230,418** | **−11.44%** | 382,896 → **350,979** | **−8.34%** | 3.629 → 3.506 |
| `ConditionalFormattingSamples.xls` | file | 1,571,067 → **1,411,127** | −10.18% | 473,603 → **448,167** | −5.37% | 3.317 → 3.149 |
| `54016.xls` | owned | 5,978,608 → 5,904,626 | −1.24% | 1,192,337 → **1,118,523** | **−6.19%** | 5.014 → 5.279 |
| `54016.xls` | file | 6,116,444 → 6,042,242 | −1.21% | 1,256,486 → 1,188,807 | −5.39% | 4.868 → 5.083 |
| `WithCustomViews.xls` | owned | 628,081 → 620,653 | −1.18% | 149,274 → 146,301 | −1.99% | 4.208 → 4.242 |
| `WithCustomViews.xls` | file | 708,942 → 701,921 | −0.99% | 190,589 → 185,277 | −2.79% | 3.720 → 3.788 |

The before leg reproduces change [0576](0576-xls-sst-scan-without-materialization.md)'s
*after* leg — 1,389,409 against its 1,391,409 on the flagship, 620,653's
predecessor 628,081 against its 627,989, 5,978,608 against its 5,979,090 — to
within 0.2%, 0.02% and 0.01%. That is the control that this record's before leg
is the branch tip.

**Cycles fall by more than instructions on two of the three fixtures**, and that
is the finding. The removed work is a *pointer chase*: `sector = fat[sector]`,
repeated, each load dependent on the last, with nothing for the core to overlap.
It is cheap in instructions — 24 per link — and expensive in cycles.
`54016.xls` loses 1.24% of its instructions and 6.19% of its cycles.

### Wall clock, paired medians, A/B/B/A, two windows

`p50` of the measured operation, nanoseconds, 100 warmups and 3,000 samples.
`dir1` is the first before→after pair and `dir2` the second; `A/A` and `B/B` are
the **same binary** against itself across the same window and are the noise
floor.

| fixture | source | before p50 | after p50 | dir1 | dir2 | A/A | B/B |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | owned | 79,760 | **72,956** | **−8.46%** | **−8.61%** | +0.15% | −0.01% |
| `ConditionalFormattingSamples.xls` | file | 98,360 | **90,925** | **−7.38%** | **−7.74%** | −0.16% | −0.56% |
| `54016.xls` | owned | 261,861 | **248,436** | **−5.05%** | **−5.21%** | +0.09% | −0.08% |
| `54016.xls` | file | 275,582 | **265,444** | −3.72% | −3.64% | −0.04% | +0.04% |
| `WithCustomViews.xls` | owned | 32,130 | **30,820** | −4.33% | −3.83% | −0.19% | +0.32% |
| `WithCustomViews.xls` | file | 39,896 | **38,520** | −3.36% | −3.54% | +0.03% | −0.16% |

A second window, captured separately, agrees: −8.62% / −8.52% and −7.69% /
−6.06% on the flagship, −5.35% / −5.36% and −3.94% / −3.89% on `54016.xls`,
−2.79% / −2.60% and −3.28% / −2.96% on `WithCustomViews.xls`.

**The noise floor, measured in these windows.** Across all 24 same-binary cells
the widest excursion is **0.94%**. Every improvement is 2.7 to 9.2 times that,
and the two directions agree to within 1.7 percentage points in every cell of
both windows. p99 moves with the medians rather than against them: −7.36% /
−6.56% on the flagship, −5.16% / −4.07% on `54016.xls`.

**Change 0574 predicted about 4.5 µs on the flagship and "0.20% to 0.26% and
worth nothing" elsewhere. The first half is beaten and the second half is
falsified.** The flagship saves 6.8 µs, and `54016.xls` — for which 0574
predicted 0.26% from its instruction share — saves **5.2%**. The prediction was
made in instructions, and instructions are the wrong currency for a dependent
load chain. That mis-ranking is the same one that sank the first implementation.

### The rejected first implementation

Change 0574 named the mechanism: give the globals scan a retained
`SharedOleStreamCursor`, the primitive change 0566 built for the worksheet scan.
That was built, and it works — but it is slower. Measured against the same before
binary, from the same worktree:

| leg | flagship chain links | instructions | cycles | IPC | flagship owned p50 |
| --- | ---: | ---: | ---: | ---: | ---: |
| before | 5,796 | 1,389,409 | 382,896 | 3.629 | 79,805 |
| **retained cursor — rejected** | **1,076** | 1,282,249 (−7.7%) | **397,030 (+3.7%)** | 3.230 | **+3.23% / +3.93%** |
| **resumable hint — retained** | 2,099 | **1,230,418 (−11.4%)** | **350,979 (−8.3%)** | 3.506 | **−8.46% / −8.61%** |

The cursor form **removes more chain links than the change that landed** — 1,076
against 2,099, because a cursor's read loop *is* the chain walk and never walks a
prefix separately — and is still worse on every cell in both directions:
+3.2%/+3.9% on the flagship, +2.3%/+2.3% on `WithCustomViews.xls`, +0.5%/+0.7% on
`54016.xls`, against a same-binary floor of 0.84% in that window.

The reason is visible in IPC, which falls from 3.629 to 3.230. A cursor advances
through four `Result`-returning methods per sector — `read_exact`,
`normalize_state`, `physical_span`, `advance_within` — each matching on the
`Fat`/`MiniFAT`/`End` state and each feeding the next, where `read_stream_range`
has one flat loop whose only serial dependency is the FAT load itself. Its added
instructions are modest (about 28 per sector) and its added *latency* is not.
`#[inline]` on the three helpers was tried in a third binary and recovered well
under a quarter of the flagship regression, which is why the cost is attributed
to the per-sector dependency chain rather than to call overhead. That capture is
**not retained** and no figure from it is cited here; it is reported as a
qualitative check that was run, not as evidence.

The evidence and the exact rejected source are retained in
[`rejected-cursor/`](results/change-0579/rejected-cursor/README.md), because the
next reader of change 0574's opportunity 3 will otherwise implement it the way
that record proposed.

## Validation preserved

Nothing about which inputs are refused, when, or with what text has changed.
Because `read_stream_range` is now `read_stream_range_hinted` with a fresh hint,
this is structural rather than a claim about two parallel implementations: there
is only one loop, one set of bound checks, one set of refusals and one fence.

### The links a hint skips are links it already validated

A hint is only ever recorded at an ordinal the walk has just reached through
`next_chain_sector`, having passed that function's marker check, index check and
successor-marker check, and having been rejected if the chain terminated early.
Resuming from it skips exactly those links and walks — and checks — every link
ahead of it. `"FAT chain ends before stream range"` and
`"MiniFAT chain ends before stream range"` therefore still fire at the same
ordinal with the same text, and `next_chain_sector`'s
`"invalid sector marker 0x%08X in FAT chain"` and
`"invalid sector index N in FAT"` are untouched, still raised from the same
function.

`"Sector N is outside the file"` is unchanged in both readers: the hint does not
reach `read_sector_run` or `flush_pending_range`, which is where that check lives
and which still name the same N. The truncated-final-sector handling, the
declared-length checks, cycle detection in `collect_exact` at open, and the
whole-container ownership and overlap validation are all untouched — this change
edits neither `file.rs` nor any open-time path.

### Three independent properties stop a hint from positioning the wrong read

This is the failure the change had to make impossible, so it is closed three
ways and each way is tested.

1. **A hint cannot cross readers.** `StreamChainHint<'a>` borrows the
   `SharedOleFile` it came from and is obtainable only from that reader's
   `chain_hint()`. It therefore cannot outlive its reader, no second live reader
   can occupy its address, and a hint offered to a different reader is discarded
   on a pointer comparison. The borrow is what makes this a property rather than
   a convention: a hint cannot even be constructed without a reader.
2. **A hint cannot cross streams.** It records the directory entry SID, which
   allocation table it indexes, and the stream's first sector, and is discarded
   unless all three match.
3. **A hint cannot move a read backwards.** A CFB chain is singly linked, so a
   hint sitting *past* the requested ordinal cannot serve it. Monotonicity is
   checked, not assumed: such a hint is discarded and the walk restarts at the
   stream's first sector.

A discarded hint costs exactly what an unhinted read costs — the walk starts at
sector zero, which is the pre-change behaviour. **A hint can only remove work it
has already done; it can never redirect a read.**

The caller obeys the same discipline without relying on it. `GlobalsBuffer::ensure`
returns early for `need <= filled` and every fill starts at `filled`, so the scan
is monotonic by construction; properties 2 and 3 are the reader's defence against
a future caller that is not.

### Policy

No `unsafe`. `crates/litchi-cfb/src/lib.rs` and `crates/litchi-xls/src/lib.rs`
keep their existing policies; this change adds no `unsafe` block in production or
test code. Fallible allocation is unaffected — the change allocates nothing at
all. No ADR-governed boundary, no dependency edge and no error type changes. The
public surface grows by one `Copy` type and two methods on an existing type.

## Correctness evidence

### Tests that fail against the pre-change code

**None, and that is the honest statement.** `read_stream_range_hinted`,
`chain_hint` and `GlobalsBuffer`'s hint field do not exist in the pre-change
tree, so the ten new tests cannot compile there, let alone fail. This change is
behaviour-preserving by construction and no behavioural test can flip. Unlike
change 0576, it also moves no *resource* counter that a test can assert on:
allocations, reads, bytes and source-version observations are all identical by
design, and a chain step is not observable from outside `litchi-cfb`.

The new tests are therefore justified by what they catch in the **new** code, and
by the corpus differential and the counter cells above, which are controls that
must pass on both sides.

### Mutation checks

Six mutations were applied to the merged tree and the suite re-run.

| mutation | caught by |
| --- | --- |
| M1: trust a hint that came from another reader | `a_hint_bound_to_another_reader_is_ignored` |
| M2: trust a hint sitting past the requested ordinal | `a_hint_past_the_requested_offset_is_ignored` |
| M3: drop the SID from the stream identity | the crate's own dead-field lint; individually redundant (see below) |
| M4: drop the allocation table from the stream identity | the same |
| M5: record the resumed position one ordinal early | `a_hinted_fat_read_matches_an_unhinted_one_over_every_partition`, `a_hinted_read_of_a_fragmented_fat_chain_matches_an_unhinted_one`, `a_hint_costs_one_chain_walk_where_an_unhinted_read_costs_one_per_call` |
| M6: drop the first sector from the stream identity | the crate's own dead-field lint; individually redundant |

**M1 initially survived, and that found a real gap.** The first version of the
cross-reader test used two readers over the *same bytes*, so the foreign hint
happened to name the correct sector and the mutation produced the right answer
for the wrong reason. The test now uses two files that lay the **same logical
stream on different physical sectors** — `sample_bytes` and its
`fragmented_large_bytes` variant, which the test asserts agree in SID, allocation
table and first sector, so reader identity is the only thing separating them —
and it asserts the control leg reads the fragmented file's own bytes before
comparing. Without that, M1 would have let a foreign hint position a read, which
is precisely the silent contract break this change had to avoid.

**M3, M4 and M6 do not compile**: removing any one of the three stream-identity
comparisons makes its `ChainResume` field dead and the crate denies unused code.
They are also each *individually* redundant given the others, and this is stated
rather than hidden: a valid CFB cannot give two live streams the same first
sector — `validate_stream_allocations` refuses overlapping claims at open — and a
directory entry is FAT-resident or MiniFAT-resident, never both. The three fields
are defence in depth against a future caller, and the reachable discrimination —
one hint used across two streams of one file — is exercised by
`a_hint_taken_from_another_stream_is_ignored`.

### The differential

`assert_hinted_matches_unhinted` reads one stream twice over the same ascending,
non-overlapping partition — once with a plain `read_stream_range` per piece, once
with one hint carried across all of them — and asserts the two legs are
indistinguishable from outside the reader: same bytes, **same positional reads in
the same order**, same number of source-version observations. It is driven over:

| shape | partitions |
| --- | --- |
| contiguous FAT chain, 8 KiB | 1, 7, 511, 512, 513, 1000, 4096 and 8192-byte pieces, plus the doubling window the globals scan actually issues |
| **fragmented** FAT chain, rewired non-monotonically with a patterned payload so a leg reading sectors in the wrong order cannot compare equal | 1, 512, 700, 1024, plus the doubling window |
| MiniFAT-resident stream, 4,095 bytes | 1, 63, 64, 65, 512, 1000, plus the doubling window |
| 4,096-byte sector size | 1, 64, 999, 4000 |

`a_hint_does_not_change_a_refusal_or_its_text` reads a range through a warm hint
and then compares the refusal text of a range past the stream end, an overflowing
offset, a missing stream and an empty path against the unhinted reader's, string
for string.

At the scan level,
`the_globals_scan_reads_identically_with_and_without_its_chain_hint` drives
`GlobalsBuffer` through a fixed ascending fill schedule on a real fixture and
compares its read ranges, bytes and source-version count against the same
extents issued through plain `read_stream_range` — where the comparison leg
recomputes the extents independently, so it shares no code with the leg under
test.

### Suites

| slice | result |
| --- | --- |
| `litchi-xls`, `litchi-cfb`, `litchi-biff`, `litchi-ole-common`, `litchi-doc`, `litchi-ppt`, `litchi-xlsx`, `litchi-xlsb`, all features | **6,355 tests, 0 failures**, 233 binaries |
| `cargo fmt --check -p litchi-cfb -p litchi-xls` | clean |
| `cargo clippy -p litchi-cfb -p litchi-xls --all-features --all-targets -- -D warnings` | clean |
| `cargo doc -p litchi-cfb -p litchi-xls --no-deps` | clean |
| `python3 -B tools/non_iwork_gate.py clippy` | clean |
| `litchi` facade library tests | 380 pass, **3 fail** |

The three facade failures are
`document::doc::tests::{filesystem_odt_keeps_source_owner_with_malformed_ooxml_catalog,
owned_odt_bytes_keep_odt_owner_with_malformed_ooxml_catalog,
managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal}`. They are
DOCX/ODT/OPC failures, **pre-existing at HEAD**, and were already recorded as
such by change 0576. `tools/check_example_targets.py` still reports duplicate
example targets in the iWork crates, also pre-existing and outside this change's
scope.

## Limitations

No cold-cache, physical-device, remote or range-source, peak-RSS,
allocation-profile, concurrency-scaling, real-producer or cross-platform result
is claimed. The latency figures are warm-cache, page-cached, on a staged
immutable copy, on one host, from two binaries that differ in three files, and
**host quiescence is not established**: other agents were active on the machine
throughout. The A/A and B/B columns and the second window are the defence against
that and are reported for every cell.

**Only XLS was measured.** `litchi-doc` calls `read_stream_range` in loops over
`WordDocument` and the table streams at increasing offsets — the same shape that
makes this walk quadratic — and is neither changed nor measured here. It can
adopt the hint without any change to `litchi-cfb`, and sizing that is a separate
change.

**The saving is concentrated and the corpus understates it.** It scales with the
number of fills a scan issues over the length of the chain, so it is worth 8.5%
on a 1.4 MB workbook and nothing at all on a small one: `SimpleMultiCell.xls`
walks 4 links before and after. Two of the eight profiled fixtures move by less
than 0.15% of instructions, in the *positive* direction, which is the hint's own
bookkeeping on files with nothing to save. This is the corpus shape change 0570
described: the right corpus for asking how often this helps, the wrong one for
asking how much.

**The worksheet scan is untouched.** It uses a `SharedOleStreamCursor` (change
0566) and keeps it. The cursor measurement in this record was taken on the
*globals* path and does not establish anything about the worksheet path, where
the cursor replaced repeated `read_stream_range` calls whose prefix walks began
deep in the Workbook stream and were far longer. Whether that path would also be
faster on a hint is an open question this record does not answer.

**The read loop's own chain walk remains.** By difference against the cursor
leg's 1,076, about 1,023 of the flagship's 2,099 remaining links are run
discovery inside the requested ranges. That work is the read, not a prefix, and
removing it would mean not reading the bytes. The callgrind profile does not
separate the two loops directly — they are one symbol — so that split is a
derivation from two measurements, not a third measurement.

Callgrind instruction counts carry change 0574's caveat unchanged: ERMS string
loops are instrumented per iteration, so `memset` and `memcpy` shares are upper
bounds. This change touches neither, and the symbols that moved are not string
loops. Callgrind's flagship figure falls 6.22% where `perf` measures 11.44% for
the same change; that is the same denominator effect and is reported rather than
reconciled. In absolute terms the two agree: 157,901 against 158,991
instructions removed, within 0.7%.

[Captures, profiles, the corpus differential, the rejected implementation and the
replay script](results/change-0579/README.md).
