# 0627: the OLE2 readers get a range source, and the first price list for one

Status: retained, harness-only. Fourteen opt-in selectors and a first
descriptive baseline of the source-backed XLS and PPT read scenarios over a
simulated range source, with an owned-source control for every phase.
`performance_claim: none` — this record carries deterministic request counts,
request-sequence identity digests and paired medians with a measured A/A floor,
not a claim-registry entry. **No file under `crates/` was modified.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is evidence gap 4 of change
[0587](0587-remaining-opportunity-survey.md) and the first of that record's
"cheapest high-value additions". The survey's table states the gap in one line:
three independent OOXML range-source mechanisms exist — `SimulatedRangeSource`
over OPC and XLSX, a dedicated `pptx_range_source` module, and provider pacing
(0447/0448) — and **no** OLE2 one; not a single CFB, XLS, DOC or PPT case had
ever run over a caller-supplied remote source. It also states why the gap is
cheap to close: *"the CFB readers already take `ReadAt`, so this is harness
plumbing"*.

It is harness plumbing. The consequence is not.
`docs/GOAL.md` names caller-supplied remote and range sources as a benchmarked
dimension for both priority formats, and until this record the program could
rank an OLE2 change on instructions, cycles, bytes and allocations but had no
way to say what it cost a caller who pays per request. Change 0587 ranks
**XLS-6** on exactly that axis — "about 30% of the flagship open as an
instruction upper bound … at a cost of about 46 more requests, a range-source
regression" — and had to call the request cost *modelled*, because nothing in
the repository could measure it.

## What was changed

Two files under `tools/`, plus documentation:

* `tools/perf-baseline/src/ole2_range_source.rs` (new, 1,113 lines including
  its five unit tests): the corpus builders, the fourteen selectors' runners,
  the two gates and the evidence block.
* `tools/perf-baseline/src/lib.rs`: the module declaration, fourteen `Case`
  variants with their `name`/`parse_case` mappings, the
  `is_ole2_range_source`/`ole2_range_source_plan` classifiers, the exclusion
  from the generic corpus loop, the explicit refusal in
  `run_case_with_config`, the `--ole2-file` option, one `SourceSummary` field,
  and the registry count.
* `tools/test_perf_baseline_source_policy.py`, `tools/perf-baseline/README.md`
  and `docs/performance/CORPUS_MANIFEST_V2.md`.

The tracked diff is +402 / -1 across those four files, plus the new module and
this record's packet.

### The selectors

Seven run over `SimulatedRangeSource` — the same simulator the OPC and XLSX
range-source cases use, which splits every logical read into physical requests
no larger than the configured maximum and pays a deterministic fixed latency,
per-request overhead and bandwidth for each one. Seven are owned-source
controls: the *same* phases over the *same* bytes through a plain instrumented
in-process `ReadAt`.

```text
xls_range_source_open                       xls_owned_source_control_open
xls_range_source_open_list_worksheets       xls_owned_source_control_open_list_worksheets
xls_range_source_open_one_cell              xls_owned_source_control_open_one_cell
xls_range_source_open_all_cells             xls_owned_source_control_open_all_cells
xls_range_source_open_full_text             xls_owned_source_control_open_full_text
ppt_range_source_open                       ppt_owned_source_control_open
ppt_range_source_open_one_shape_text        ppt_owned_source_control_open_one_shape_text
```

The XLS five are the `open`, `list`, `one-cell` scenarios of the existing
source-backed family plus the `all-cells` and `full-text` scenarios change
[0605](0605-xls-retained-sheet-index.md) added; the PPT two are
`text_edit::SourceSnapshot::open` and `read_text`. Registration mirrors change
[0601](0601-perf-harness-real-producer-shape.md)'s producer-shape family
exactly — enum, `name`, `parse_case`, classifier, generic-loop exclusion, a
family block in `run()`, an explicit refusal in the generic dispatcher, usage
text, README and the corpus-manifest document.

### The corpus

A caller-named real OLE2 file supplied with `--ole2-file PATH`. The flag is
repeatable and each file is bound to its family by its **CFB stream
inventory** — a `Workbook`/`Book` stream selects XLS, a `PowerPoint Document`
stream selects PPT — so one run measures an XLS and a PPT fixture together and
no extension heuristic is involved. This is the second input in this harness
whose bytes do not come from the process, so it is treated the way 0601 treats
`--real-file`: bounded to 32 MiB, absent from `Case::DEFAULT`, and carrying its
path, size and SHA-256 in the corpus identity, together with the CFB stream
count, the sector size and the target stream's name, length and SHA-256.

Nothing about the file is assumed. The worksheet names, the selected worksheet
(the first whose walk reports a cell with a value), the selected cell (the
median valued position of that walk), the selected slide and shape (the first
position the eager reader reports with text), and every scenario's outcome are
derived from the bytes before any sample runs.

### Shape: a complete lifecycle per phase

Every selector opens a fresh source and a fresh owner inside the timed region,
then performs its operation. That is the shape the existing
`xls_source_backed_open_list_worksheets` family and the standalone
`xls_source_attribution` profiler already use, and it is the right one here:
a caller paying per request cares about the cost of *reaching* a cell, not
about a marginal query on a workbook someone else already opened. It also
keeps every phase's request sequence non-empty, which the XLSX
`xlsx_range_source_list_sheets` case deliberately does not (it asserts zero
timed requests, because an XLSX listing is served from the already-parsed
workbook part).

### A refusal is an outcome

Two of the four fixtures refuse a scenario, and the refusal string is frozen as
that scenario's oracle rather than a reason to drop the case, exactly as
`xls_source_attribution`'s full-text scenario already does. A caller paying per
request still pays for the bytes the reader read before refusing, and that
price is the thing being measured. See *Two refusals worth recording*.

## Why it is sound

**No production crate is touched.** The module reads through the public
`litchi_xls::SourceBackedWorkbook`, `litchi_ppt::text_edit::SourceSnapshot` and
`litchi_cfb::OleFile` APIs and through `litchi_core::ReadAt`, which
`SimulatedRangeSource` already implements. No `unsafe`, no
`#[global_allocator]`, no `std::env`, no `Command`; a source-policy test
asserts all four, mirroring the one 0601 added for `producer_shape`.

**No existing selector, corpus identity or default-matrix row changes.** All
fourteen are absent from `Case::DEFAULT`, whose length stays 41. The
selectable registry grows from 457 to 471 names. The schema-2 default catalog
is generated from `results/perf-regression-default-manifest-v1.json`, which
covers exactly `Case::DEFAULT`, so a case that is not in `Case::DEFAULT` never
reaches it: `test_corpus_manifest_v2` and `test_crud_coverage_index` both pass,
which is the mechanical statement that the checked `catalog_sha256` and
`content_set_sha256` did not move. A third source-policy test asserts that none
of the fourteen variant names appears inside the `DEFAULT` array.

**Bounded resources.** One 32 MiB ceiling per named file, checked from
`fs::metadata` before the single `fs::read` in the module (the policy test
asserts there is exactly one). The retained request sequence is a digest plus a
32-triple preview, not tens of thousands of triples per sample.

**Validation does not mutate, and no refusal moved.** Every scenario calls a
read-only API. The two typed refusals below are the readers' own, reported
verbatim and unchanged; no gate was relaxed to make a fixture measurable.

**ADR reading.** ADR 0003's bounded-resource rule is satisfied by the file
ceiling and the bounded evidence retention; ADR 0005's lazy-payload contract is
observed, not altered — this record measures when the existing readers read,
and changes nothing about it; ADR 0006 is untouched because no execution
context, worker pool or ambient I/O is introduced. The simulator runs on the
calling thread and sleeps on it.

### Two gates, both taken from 0572's frozen set

Change [0572](0572-ooxml-range-source-attribution.md) froze five gates before
capture; two of them are properties of a single selector rather than of a
132-arm matrix, and those two are implemented here as hard failures:

1. **Observation identity.** Every retained sample must reproduce the oracle
   derived from the file before the run. The per-sample digests are retained.
2. **Request-sequence identity.** The complete ordered
   `(offset, requested, returned)` sequence is hashed per retained sample and
   every sample must produce the same digest. 0572's wording is the property
   being checked: *"the request sequence is a pure function of the fixture,
   scenario and policy."*

Both held on every case of every fixture, over 50 retained samples each.

The cross-transport half of 0572's determinism gate is checked in the analysis
rather than inside one selector, because the two transports are two separate
selectors here: for every fixture and scenario the range-source leg's logical
read calls, logical read bytes and observation are asserted equal to the
owned-source leg's by `scripts/summarize.py`, which fails the summary
otherwise. They were equal everywhere.

### One model difference from 0572, stated

0572's probe models a capped request as a **short read** and lets the caller
loop; this harness's simulator **loops the cap itself** inside one `read_at`.
The physical request counts agree; the call boundaries do not. 0572's model
also has no per-request overhead term at all, so its delayed arm is reproduced
here as `--range-request-overhead-us 0`. The transport used throughout is
0572's delayed arm expressed in this harness's four parameters: **1 ms of fixed
service per request, no overhead term, 100 MiB/s, 64 KiB maximum physical
range**.

## Measured

Base commit `344ed0298fdabe42c2596a4309db8c7a11eaf355`, branch
`perf/0627-ole2-range-source-selectors`, one release binary
(`--release --locked`, SHA-256 `d40979be2b58a182…8728fd1`), CPU 20 via
`taskset`, 20 warm-ups and 50 retained samples per case, all legs in one
window with seven other agents active on the host. Fixtures:
`test-data/ole/xls/WithCustomViews.xls` (165,888 bytes),
`test-data/ole/xls/ConditionalFormattingSamples.xls` (1,402,368),
`test-data/poi/test-data/spreadsheet/54016.xls` (984,576),
`test-data/poi/test-data/slideshow/45543.ppt` (385,024).

### Deterministic counts

Every figure below was identical on all 50 retained samples of both
transports; `scripts/summarize.py` fails if the two legs disagree on logical
read calls, logical read bytes or the observation, and they never did.

| fixture | scenario | logical reads | logical bytes | physical requests | physical bytes | sequence digest stable | observation |
|---|---|---:|---:|---:|---:|---|---|
| `WithCustomViews.xls` | open | 16 | 110,242 | 16 | 110,242 | yes | `worksheets:3` |
| `WithCustomViews.xls` | open+list-worksheets | 16 | 110,242 | 16 | 110,242 | yes | `names:3:0bfdc4a92c188add30ce1c570493d689775...` |
| `WithCustomViews.xls` | open+one-cell | 24 | 150,818 | 24 | 150,818 | yes | `string:25:Produto &\nNovos Negócios` |
| `WithCustomViews.xls` | open+all-cells | 896 | 257,066 | 896 | 257,066 | yes | `cells:3325:804257a9e49ac1cde6059245bf513010...` |
| `WithCustomViews.xls` | open+full-text | 898 | 257,894 | 898 | 257,894 | yes | `text:107420:bdd9a79ef483a767ac8cd118a46ca2a...` |
| `ConditionalFormattingSamples.xls` | open | 53 | 565,201 | 53 | 565,201 | yes | `worksheets:16` |
| `ConditionalFormattingSamples.xls` | open+list-worksheets | 53 | 565,201 | 53 | 565,201 | yes | `names:16:9edfc48f3d76d069f3b1d28bd979afa02b...` |
| `ConditionalFormattingSamples.xls` | open+one-cell | 62 | 611,237 | 62 | 611,237 | yes | `string:8:Icon set` |
| `ConditionalFormattingSamples.xls` | open+all-cells | 97 | 612,624 | 97 | 612,624 | yes | `cells:61:661c0620faadfe8b30d7c1c9f399af960a...` |
| `ConditionalFormattingSamples.xls` | open+full-text | 394 | 790,147 | 394 | 790,147 | yes | `refused:source XLS parse error: Invalid rec...` |
| `54016.xls` | open | 40 | 317,171 | 40 | 317,171 | yes | `worksheets:1` |
| `54016.xls` | open+list-worksheets | 40 | 317,171 | 40 | 317,171 | yes | `names:1:0e50f35f1dd82c89a24071766ea8be3af3f...` |
| `54016.xls` | open+one-cell | 66 | 933,004 | 66 | 933,004 | yes | `string:8:131103G2` |
| `54016.xls` | open+all-cells | 16,145 | 1,256,139 | 16,145 | 1,256,139 | yes | `cells:38950:cc68b88da7bdea28644aeb4801cb41f...` |
| `54016.xls` | open+full-text | 16,145 | 1,256,139 | 16,145 | 1,256,139 | yes | `text:840804:d88d103e72b84b430a75ed4c0da59cc...` |
| `45543.ppt` | open | 8 | 779,264 | 18 | 779,264 | yes | `slides:11` |
| `45543.ppt` | open+one-shape-text | 73 | 792,661 | 83 | 792,661 | yes | `refused:selected PPT shape text has an unsu...` |

### What the counts say

Four things stand out, all of them new facts about these readers:

* **An XLS open is not a small read.** It costs 16 requests and 110,242 bytes
  on `WithCustomViews.xls` (66.5% of the whole file), 53 and 565,201 on
  `ConditionalFormattingSamples.xls` (40.3%), 40 and 317,171 on `54016.xls`
  (32.2%). Listing the worksheets after that open is **free**: the list
  scenario's counts are identical to the open's on all three fixtures, because
  the names come from metadata the open already parsed.
* **The PPT open reads the file about twice.** 8 logical reads returning
  779,264 bytes from a 385,024-byte file — 2.02x. `SourceSnapshot::open`
  validates the complete CFB index and then takes a complete-artifact
  fingerprint, and on this fixture that is two passes over the bytes. It is
  also the only scenario here where the 64 KiB cap bites: 8 logical reads
  become 18 physical requests.
* **Change 0605's whole-sheet walk is the single most expensive thing a
  range-source caller can ask an XLS reader for.** `54016.xls` all-cells is
  16,145 requests and 1,256,139 bytes — 1.28x the file — against 66 requests
  for one cell. At this transport that is 16.2 seconds of modelled service for
  one worksheet. The walk is still the right primitive: the alternative 0605
  replaced was one full validated scan *per cell*, which for 38,950 cells
  would be roughly 38,950 x 26 requests.
* **Full text and all cells are the same read on a single-sheet workbook.**
  `54016.xls` reports exactly 16,145 requests and 1,256,139 bytes for both.
  On the three-sheet `WithCustomViews.xls` full text costs 2 requests more
  than the walk of one sheet (898 against 896).

### The request-size distribution, and where the lever is

The harness's fixed size buckets, first retained sample of each range-source
leg (every sample agreed):

| fixture | scenario | 1–512 B | 513–4,096 | 4,097–16,384 | 16,385–65,536 | > 64 KiB |
|---|---|---:|---:|---:|---:|---:|
| `WithCustomViews.xls` | open | 7 | 5 | 2 | 2 | 0 |
| `WithCustomViews.xls` | open+one-cell | 9 | 8 | 5 | 2 | 0 |
| `WithCustomViews.xls` | open+all-cells | 812 | 77 | 5 | 2 | 0 |
| `ConditionalFormattingSamples.xls` | open | 30 | 4 | 10 | 9 | 0 |
| `ConditionalFormattingSamples.xls` | open+full-text | 334 | 25 | 26 | 9 | 0 |
| `54016.xls` | open | 26 | 3 | 6 | 5 | 0 |
| `54016.xls` | open+one-cell | 29 | 14 | 8 | 15 | 0 |
| `54016.xls` | open+all-cells | 16,061 | 61 | 8 | 15 | 0 |
| `45543.ppt` | open | 2 | 4 | 0 | 12 | 0 |
| `45543.ppt` | open+one-shape-text | 65 | 5 | 1 | 12 | 0 |

**16,061 of `54016.xls`'s 16,145 all-cells requests are 512 bytes or smaller**,
and they carry a small share of the 1,256,139 bytes: the open's 26 + 3 + 6 + 5
requests move 317,171 of them. The whole-sheet walk's cost on a range source is
therefore almost entirely *request count*, not bytes — 16.1 seconds of fixed
service against 12 ms of modelled transfer at 100 MiB/s. That is the shape a
read-ahead, a buffered stream cursor, or a coalescing `ReadAt` adapter would
attack, and it is a different lever from anything the instruction-count records
have ranked, because on a warm local source those 16,061 reads are `memcpy`
calls that cost almost nothing. No such change is proposed here; the point is
that the axis now exists and can be measured.

The 64 KiB cap is invisible on every XLS leg — no bucket exceeds it and
physical requests equal logical reads throughout — and bites only on the PPT
open, where 8 logical reads become 18 physical requests.

### Two refusals worth recording

Neither is introduced here; both are the readers' own typed refusals, met for
the first time on this axis.

* `ConditionalFormattingSamples.xls` **full text is refused** — `source XLS
  parse error: Invalid record` — after 394 requests and 790,147 bytes. A
  range-source caller pays for 56.4% of the file before being told no. The
  same fixture's open, list, one-cell and all-cells scenarios all succeed, so
  this is specific to the full-text projection.
* `45543.ppt` **one-shape text is refused** — `selected PPT shape text has an
  unsupported dependency record` — after 73 logical reads and 792,661 bytes.
  The target is the first slide shape the *eager* reader reports with text, so
  the shape exists and has text; `text_edit::SourceSnapshot::read_text`
  declines it. Spot checks on four further real decks (`SampleShow.ppt`,
  `text_shapes.ppt`, `text-margins.ppt`, `ppt_with_png.ppt`) refuse the same
  way, while `litchi-ppt`'s own test for this API uses a generated fixture.
  Nothing is claimed about the corpus-wide rate: that is a refusal census, and
  it is change 0587's evidence gap 3, not this record's.

### Pricing a proposal on a range source

This is what the record is for. Change 0587 ranks **XLS-6** partly on a cost it
could only model: *"at a cost of about 46 more requests, a range-source
regression."* That sentence now has a denominator. The flagship
`ConditionalFormattingSamples.xls` open is **53 requests** with a **58.39 ms**
modelled service floor at this transport, so 46 more requests is **+46 ms of
fixed service, +78.8% on the floor and a 1.87x open** — before any extra bytes.
XLS-6's retained upside is about 30% of the flagship open as an instruction
*upper bound* and 5.1% of aggregate open time. So the same change that removes
at most 30% of an open's instructions adds 78.8% to the same open's modelled
request cost. The design record for XLS-6 can now state that trade-off in
measured units on both sides, which is what its density gate has to be argued
against.

The general form is the same for every proposal: at this transport a saved
request is worth 1 ms plus `bytes / 100 MiB/s`, and the per-scenario request
counts above are the denominators. Nothing here says what a *real* network
costs; it says what this configured model costs, and the model is the
program's agreed one.

### The A/A floor

The owned-source controls were run twice, back to back, in the same window:
the second run against the first is an A/A on this host with seven other
agents active.

| fixture | scenario | A1 p50 (us) | A2 p50 (us) | delta p50 | delta p95 | delta p99 |
|---|---|---:|---:|---:|---:|---:|
| `WithCustomViews.xls` | open | 25.28 | 25.07 | -0.83% | +0.47% | -1.89% |
| `WithCustomViews.xls` | open+list-worksheets | 25.84 | 25.45 | -1.51% | -2.39% | +6.13% |
| `WithCustomViews.xls` | open+one-cell | 69.61 | 68.67 | -1.35% | -1.82% | -0.27% |
| `WithCustomViews.xls` | open+all-cells | 1,452.05 | 1,444.04 | -0.55% | -0.57% | -1.07% |
| `WithCustomViews.xls` | open+full-text | 1,128.71 | 1,109.42 | -1.71% | -1.05% | +0.66% |
| `ConditionalFormattingSamples.xls` | open | 72.53 | 72.83 | +0.41% | -2.74% | +8.77% |
| `ConditionalFormattingSamples.xls` | open+list-worksheets | 81.93 | 81.26 | -0.82% | -0.86% | -26.33% |
| `ConditionalFormattingSamples.xls` | open+one-cell | 87.59 | 87.95 | +0.41% | +4.27% | -0.94% |
| `ConditionalFormattingSamples.xls` | open+all-cells | 132.65 | 132.94 | +0.22% | +0.59% | +1.20% |
| `ConditionalFormattingSamples.xls` | open+full-text | 417.42 | 415.76 | -0.40% | -0.46% | -0.97% |
| `54016.xls` | open | 191.45 | 191.96 | +0.27% | -0.37% | -1.50% |
| `54016.xls` | open+list-worksheets | 191.97 | 192.45 | +0.25% | +0.35% | +0.14% |
| `54016.xls` | open+one-cell | 822.29 | 821.41 | -0.11% | +0.30% | +0.85% |
| `54016.xls` | open+all-cells | 15,973.77 | 15,923.36 | -0.32% | -0.59% | -1.41% |
| `54016.xls` | open+full-text | 18,791.50 | 18,883.97 | +0.49% | +1.79% | -6.00% |
| `45543.ppt` | open | 410.98 | 412.13 | +0.28% | +1.28% | +3.05% |
| `45543.ppt` | open+one-shape-text | 447.00 | 454.54 | +1.69% | +0.23% | +0.58% |

**Worst |A/A| at p50: 1.71%**, and 17 of 17 rows are inside 1.71%. That is
well inside the host's stated floor (about 4% at p50), so the owned medians
below are readable at the tens-of-percent scale, not at the few-percent scale.
The p99 column is noisier, as always: one row moves 26%.

### Range source against owned source

The ratio column is the honest shape of the result and the reason the elapsed
tier is *modelled*: the last column shows that **86.7% to 94.8%** of the range
leg's median is the transport model's own arithmetic. These rows say what the
configured transport costs, not what a network costs.

| fixture | scenario | owned p50 (ms) | range p50 (ms) | ratio | modelled service floor (ms) | floor share of range p50 |
|---|---|---:|---:|---:|---:|---:|
| `WithCustomViews.xls` | open | 0.025 | 17.996 | 711.8x | 17.051 | 94.8% |
| `WithCustomViews.xls` | open+list-worksheets | 0.026 | 17.995 | 696.3x | 17.051 | 94.8% |
| `WithCustomViews.xls` | open+one-cell | 0.070 | 26.879 | 386.1x | 25.438 | 94.6% |
| `WithCustomViews.xls` | open+all-cells | 1.452 | 950.817 | 654.8x | 898.452 | 94.5% |
| `WithCustomViews.xls` | open+full-text | 1.129 | 953.082 | 844.4x | 900.460 | 94.5% |
| `ConditionalFormattingSamples.xls` | open | 0.073 | 61.578 | 848.9x | 58.390 | 94.8% |
| `ConditionalFormattingSamples.xls` | open+list-worksheets | 0.082 | 61.735 | 753.5x | 58.390 | 94.6% |
| `ConditionalFormattingSamples.xls` | open+one-cell | 0.088 | 72.223 | 824.5x | 67.829 | 93.9% |
| `ConditionalFormattingSamples.xls` | open+all-cells | 0.133 | 108.670 | 819.2x | 102.842 | 94.6% |
| `ConditionalFormattingSamples.xls` | open+full-text | 0.417 | 451.022 | 1,080.5x | 401.536 | 89.0% |
| `54016.xls` | open | 0.191 | 48.584 | 253.8x | 43.025 | 88.6% |
| `54016.xls` | open+list-worksheets | 0.192 | 48.350 | 251.9x | 43.025 | 89.0% |
| `54016.xls` | open+one-cell | 0.822 | 86.341 | 105.0x | 74.898 | 86.7% |
| `54016.xls` | open+all-cells | 15.974 | 17,067.569 | 1,068.5x | 16,156.986 | 94.7% |
| `54016.xls` | open+full-text | 18.792 | 17,077.236 | 908.8x | 16,156.986 | 94.6% |
| `45543.ppt` | open | 0.411 | 26.871 | 65.4x | 25.432 | 94.6% |
| `45543.ppt` | open+one-shape-text | 0.447 | 95.720 | 214.1x | 90.559 | 94.6% |

## Correctness evidence

Five new unit tests in `ole2_range_source::tests`, all passing:
`scenarios_and_transports_describe_themselves`;
`every_ole2_range_source_case_is_opt_in_and_round_trips` (all fourteen
round-trip through `Case::name`/`parse_case`, all fourteen answer
`is_ole2_range_source`, none is in `Case::DEFAULT`);
`xls_corpus_derives_its_targets_from_the_file`;
`xls_range_source_and_owned_legs_agree_on_logical_reads` (the two transports
agree on logical read calls, logical read bytes and the observation; the
range leg's physical requests are at least its logical reads; the owned leg
reports no physical requests);
`ppt_corpus_selects_a_text_shape_and_freezes_its_outcome`.

Three new source-policy tests in `tools/test_perf_baseline_source_policy.py`,
mirroring the three change 0601 added: the module declares no `unsafe`, no
`#[global_allocator]`, no `std::env` and no `Command`; `--ole2-file` is
bounded, self-identifying and performs exactly one `fs::read`; and none of the
fourteen variants appears in the `DEFAULT` array.

Gates, all in the worktree (only `tools/perf-baseline` was touched, so only
that package is gated):

| Gate | Result |
| --- | --- |
| `cargo fmt --all -- --check` | clean |
| `cargo clippy --release --locked --all-targets` | clean, 0 warnings (workspace lints are deny) |
| `cargo test --release --locked` | 513 passed, 0 failed, 1 ignored in the library; 12 + 2 + 4 passed across the binary targets; 0 failed anywhere |
| `cargo doc --release --locked --no-deps` | clean (rustdoc lints are deny) |
| `python3 -m unittest` on all 26 `tools/test_*.py` modules | 23 OK, 3 pre-existing failures |
| `tools/check_perf_claims.py --mode strict` | OK, 10 claims validated |
| `tools/check_report_claim_classification.py` | OK, 167 REPORT rows across 2 tables |
| `tools/validate_crud_coverage_index.py` | OK, 15 categories, 33 mapped selectors |
| `tools/non_iwork_gate.py verify` | OK, 45 bulk tree roots, 35 facade safe trees, 1 combined tree |

The three failing Python modules were reproduced with this change's tracked
edits stashed at the base commit, so they are pre-existing and unrelated:
`test_check_crate_boundaries` (an iWork crate-boundary edge count, 241 against
an asserted 240 — iWork is excluded from this programme and untouched here),
`test_native_odf_resave` (three tests read a LibreOffice filter registry that
is not installed on this host), and `test_perf_claims`
(`test_seed_registry_is_structurally_valid`, a claim-registry literal
mismatch). `test_corpus_manifest_v2` and `test_crud_coverage_index` both pass,
which is the mechanical statement that the checked default catalog SHA-256 did
not move. The last four gates are unchanged-state checks — this change
registers no claim, adds no REPORT claim row and maps no new selector into the
coverage index — and are run to prove that.

The two per-case gates held on every leg. The baseline ran 51 legs — 17
owned-source, the same 17 repeated as the A/A, and 17 range-source — and on
each one all 50 retained samples reproduced the frozen oracle; every
range-source leg produced exactly one request-sequence digest across its 50
samples. The cross-transport check in `scripts/summarize.py` passed for all 17
fixture/scenario pairs.

The gate tails are in [`results/change-0627/gates.txt`](results/change-0627/gates.txt).

## Validation preserved

No validation path was changed, relaxed or bypassed. Every scenario calls a
public read-only API; the CFB index validation, protected-container policy,
source-version check and complete-artifact fingerprint that
`SourceBackedWorkbook::from_read_at` and `SourceSnapshot::open` perform are
inside the measured region, which is why the PPT open reads more bytes than the
file contains (see below). The two typed refusals are the readers' own and are
reported verbatim. No limit was weakened: the `--ole2-file` ceiling is a new
limit, not a relaxed one.

## Limitations

**What is not claimed.**

* **No speedup, regression or production result.** Nothing under `crates/`
  changed, so there is nothing to speed up or slow down. The range-source
  numbers are a first descriptive baseline.
* **Elapsed time on the range-source leg is modelled, not measured.** It is
  sleep-driven arithmetic over a deterministic request count. Per changes 0447
  and 0448 the retained delay counters are *requested* targets, not observed
  sleeping time; OS sleep granularity adds oversleep that is never subtracted;
  and the per-call model implements no shared-link scheduler. The
  `simulated_service_floor_ns` figure is a model accounting quantity, in 0144's
  words, not observed wall-clock service.
* **No network, device, filesystem, cold-cache or page-cache result.** Every
  source is in-memory bytes. `SourceBackedWorkbook::from_path` and
  `FileSource` are not exercised, so the `fstat` per source observation that
  change 0587 identified on the `from_path` route is still unmeasured here.
* **Four fixtures, one host, one toolchain, one transport point.** The
  transport is a single (1 ms, 100 MiB/s, 64 KiB) point; nothing is claimed
  about how any of these scenarios behaves at another latency, bandwidth or
  range cap. No DOC and no CFB-level selector was added: gap 4 named four
  formats and this record closes two.
* **The owned-source control legs are CPU-bound and were taken with other
  agents active.** Their A/A floor is reported and is the only basis on which
  the owned medians should be read.
* **The all-cells and full-text scenarios walk one worksheet.** They are not a
  whole-workbook figure; the selected worksheet is named in each result's
  corpus evidence.

**Deliberately not done.** The stale count sentence at
`tools/perf-baseline/README.md:10` ("The registry has 441 selectors") predates
change 0601 and is left alone: correcting it is not this record's change, and
the authoritative count is the test-enforced `selectable_count` assertion.

## Retained evidence

[`results/change-0627/README.md`](results/change-0627/README.md) — the raw
schema-1 reports for every leg, the two scripts, the gate tails, the decision
record and the log paragraphs.
