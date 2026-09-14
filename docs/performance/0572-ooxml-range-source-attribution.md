# 0572: what a source-backed OOXML read actually asks the source for

Status: attribution only. No production change and `performance_claim: none`.
This record carries an exact, classified, ordered list of every
positional read three OOXML scenarios issue on a caller-supplied source, plus
delayed-transport wall-clock medians from a simulated transport. It confirms
six of the frozen plan's predictions, refutes five, splits a seventh, and
corrects one statement in change
[0567](0567-ooxml-single-index-per-open.md).

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## The question

Change [0561](0561-opc-repeated-positional-reads.md) concluded that collapsing
repeated positional reads is worth nothing on a warm local file and about 85% on
a latency-bearing source, and that further work here has to be justified against
a range-source measurement. Change [0567](0567-ooxml-single-index-per-open.md)
found that the XLSX streaming cell read triggers a whole-archive strict-layout
proof. **Neither had been measured on a range source.** This record measures
both.

The plan was frozen before capture at
[`results/change-0572/plan.json`](results/change-0572/plan.json) and is not
edited here. Where reality disagrees with it, reality is reported.

## Method

A counting `litchi_core::ReadAt` records every `(offset, length, returned)`
triple in issue order. It sits exactly where a caller's source sits, so nothing
is inferred: `litchi-opc`'s `SourceReader` forwards `soapberry_zip::ReaderAt`
straight to it under the `exact` policy, and through `ArchiveReadAhead` under a
`forward_start` policy. The fixture's ZIP structure is parsed independently by
the analysis script — the library is never asked where its own bytes are.

Two transports, as the plan specifies: a zero-delay control, and the 0493
configuration (1 ms fixed service per physical request, 104,857,600 bytes per
second, minimum-service combination, 65,536-byte maximum physical range served
as a short read). A third **zero-delay-plus-range-cap** arm was added beyond the
plan purely to separate a sequence change caused by the range cap from one
caused by timing; it found no difference from either, which is why it appears
only here.

Eleven fixtures x four policy-and-route combinations x three transports x five
repeats: **132 arms, 660 scenario runs.** The four combinations are `exact` on
the native leaf constructor, `exact` through the two-step adopt route, and each
`forward_start` window through the adopt route. **No read-ahead arm was run
through a native leaf constructor**, because no PPTX or XLSX leaf constructor
accepts a policy at all — see the policy gate below. The control gate is what
licenses comparing the adopt-route read-ahead arms to the native `exact`
baseline.

### The tree had to be pinned, and this matters

The first capture was taken against the working tree and **had to be thrown
away**. A large uncommitted change to `crates/soapberry-zip/src/archive.rs` —
the strict-layout proof itself, written up concurrently as change
[0573](0573-zip-single-local-header-read.md) and still uncommitted in the
working tree as this is written — was in flight, and the probe linked it. That build reported the XLSX proof as *one* read per member
(26 → 13, 264 → 132, 354 → 222 whole-scenario requests), which is 0573's
*after* state, not this record's subject. Every number below is instead from a
`git archive` of `163ac1bd67f2a0d72c27bfec60e0a2620768cc9d` extracted to a
scratch tree, so a concurrent edit under `crates/` cannot reach it.

The DOCX and PPTX counts were **identical** in both builds, which is a useful
check that the isolation works and that the contamination was confined to the
XLSX path. Its XLSX figures match the *after* figures change 0573 reports for
its own change, which is consistent with it having measured 0573 by accident —
but **that capture was discarded and is not retained here**, so nothing in this
packet establishes it, and the corroboration comes entirely from 0573's packet.
What is certain is that it would have been published as this record's baseline
if the build had not been pinned.

**The lesson is worth stating plainly for this program:** a probe crate with
path dependencies on a shared working tree measures whatever is in that tree at
link time, and gives no warning. Pin the revision.

## The shape, measured

Every scenario opens the same way, and the plan described that opening
correctly:

1. one **22-byte** end-of-central-directory read;
2. one **46-byte** first-central-record probe;
3. one read of the **whole central directory**;
4. then, per structural member, a **30-byte local header** read and a **payload**
   read — plus a **16-byte data-descriptor** read for each member that carries
   one.

| scenario | fixture | bytes | members | desc | exact | open | proof | rest |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| DOCX full text | `comment.docx` | 4,926 | 10 | 10 | **15** | 12 | — | 3 |
| DOCX full text | `endnotes.docx` | 13,553 | 16 | 0 | **13** | 11 | — | 2 |
| DOCX full text | `testComment.docx` | 65,298 | 17 | 0 | **13** | 11 | — | 2 |
| PPTX middle slide | `shape-glow-effect.pptx` | 13,691 | 17 | 0 | **19** | 17 | — | 2 |
| PPTX middle slide | `shapes.pptx` | 68,822 | 48 | 0 | **49** | 47 | — | 2 |
| PPTX middle slide | `shape-soft-edges.pptx` | 50,092 | 65 | 65 | **87** | 84 | — | 3 |
| XLSX cell A1 | `sheet-names.xlsx` | 9,425 | 13 | 0 | **40** | 11 | 26 | 3 |
| XLSX cell A1 | `universal-content.xlsx` | 7,296 | 13 | 13 | **58** | 18 | 39 | 1 |
| XLSX cell A1 | `ConditionalFormattingSamples.xlsx` | 654,688 | 132 | 0 | **354** | 89 | 264 | 1 |
| XLSX cell A1 | `SimpleNormal.xlsx` | 8,871 | 12 | 0 | **40** | 11 | 24 | 5 |
| XLSX cell A1 | `ExcelPivotTableSample.xlsx` | 19,460 | 27 | 0 | **86** | 25 | 54 | 7 |

"open" is counted by the probe itself at the constructor boundary, not inferred
from the sequence. "proof" is the strict-layout sweep described below; "rest"
is everything after it.

### What each scenario actually did

The request counts are exact, but two of the scenario labels promise more than
the fixture delivers, and that is worth stating before the tables are read.
Every fixture's outcome is retained in `results.json` as `exact_observations`:

| fixture | scenario outcome |
| --- | --- |
| `comment.docx` | **0 bytes of text**, 1 object — the body carries none |
| `endnotes.docx` | 973 bytes, 12 objects |
| `testComment.docx` | 28 bytes, 1 object |
| `shape-glow-effect.pptx` | **1 slide**, so "middle" is slide 0; 0 bytes of text |
| `shapes.pptx` | **6 slides**, middle is slide 3; 65 bytes of text |
| `shape-soft-edges.pptx` | **1 slide**, so "middle" is slide 0; 9 bytes of text |
| `sheet-names.xlsx` | A1 is `Missing` |
| `universal-content.xlsx` | refused, see below |
| `ConditionalFormattingSamples.xlsx` | refused, see below |
| `SimpleNormal.xlsx` | A1 is `Stored(Text("test"))` |
| `ExcelPivotTableSample.xlsx` | A1 is `Stored(Text("Field1"))` |

So **only `shapes.pptx` exercises a genuine middle slide**, and the DOCX
"full text" scenario extracts no text at all on one of its three fixtures. This
does not weaken the request counts — the work of locating and materializing the
part is identical whether the part yields text or not — but a reader comparing
these counts to a text-heavy corpus should know that none of these fixtures is
one.

### What "structural member" actually means

The plan predicted "three structural members for DOCX and XLSX, twenty-one for
the six-slide `shapes.pptx`". That is **wrong**, and the correct rule is
mechanical:

> the open reads `[Content_Types].xml`, **every `.rels` part in the package**,
> and the format's main part — `xl/workbook.xml` for XLSX,
> `ppt/presentation.xml` for PPTX. DOCX reads no main part during the open,
> because `word/document.xml` is the scenario's own target and falls in the read
> phase.

So the count is `1 + rels + 1`, or `1 + rels` for DOCX. It predicts **all eleven
fixtures exactly**:

| fixture | `.rels` parts | predicted | measured |
| --- | ---: | ---: | ---: |
| `comment.docx` | 2 | 3 | 3 |
| `endnotes.docx`, `testComment.docx` | 3 | 4 | 4 |
| `shape-glow-effect.pptx` | 5 | 7 | 7 |
| `shapes.pptx` | 20 | 22 | **22**, not the predicted 21 |
| `shape-soft-edges.pptx` | 25 | 27 | 27 |
| `sheet-names.xlsx`, `SimpleNormal.xlsx` | 2 | 4 | 4 |
| `universal-content.xlsx` | 3 | 5 | 5 |
| `ExcelPivotTableSample.xlsx` | 9 | 11 | 11 |
| `ConditionalFormattingSamples.xlsx` | 41 | 43 | **43** |

The open therefore scales with the **relationship-part count**, not with the
format and not with the member count. A workbook with eighteen worksheets and
eighteen drawings carries 41 `.rels` parts and pays 43 structural members, or 89
of its 354 requests, before the cell read starts. **This is the part of the
sequence the plan's model did not have**, and it is what makes two of the plan's
other predictions fail.

### DOCX and PPTX: no proof, two requests for the part

`write_text_to` costs exactly **two more requests** — a 30-byte local header and
the payload of `word/document.xml` — or **three** when the member carries a
descriptor. The middle slide costs the same two or three for
`ppt/slides/slideN.xml`. There is no strict-layout sweep in either scenario.
Both predictions hold exactly.

### XLSX: the strict-layout proof, priced

The proof is real, it is contiguous, it walks the archive in ascending
local-header order, and it precedes the first worksheet payload byte. Its cost
is **two requests per member** — a 30-byte local header, then the variable
header region (name plus extra field) — and **three** per descriptor-bearing
member.

| fixture | members | proof requests | of the cell phase | of the scenario | proof bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| `sheet-names.xlsx` | 13 | **26** | 26 of 29, 89.7% | 65.0% | 2,478 |
| `universal-content.xlsx` | 13, all descriptors | **39** | 39 of 40, 97.5% | 67.2% | 861 |
| `ConditionalFormattingSamples.xlsx` | 132 | **264** | 264 of 265, 99.6% | 74.6% | 10,740 |
| `SimpleNormal.xlsx` | 12 | **24** | 24 of 29, 82.8% | 60.0% | 2,412 |
| `ExcelPivotTableSample.xlsx` | 27 | **54** | 54 of 61, 88.5% | 62.8% | 3,405 |

The plan predicted 26, 264 and 39 for the three fixtures it named. **All three
are exactly right** — the most precise prediction in the plan, and the one
change 0567 had already established the shape of.

The plan's "more than 85 percent of the cell phase" holds on all three fixtures
it named. It does **not** hold on `SimpleNormal.xlsx` at 82.8%, one of the two
fixtures added here — that fixture's cell read also fetches shared strings, which
costs the extra requests the share is measured against.

The proof is cheap in bytes and expensive in requests: 10,740 bytes across 264
requests on the 655 KB fixture, an average of 41 bytes per request. That is the
worst possible shape for a latency-bearing source.

### An independent instrument agrees, to the byte

Change [0573](0573-zip-single-local-header-read.md), written concurrently, is
the change that collapses those two reads into one. It counts the same proof
through `soapberry-zip`'s own instrumented reader — a different instrument, at a
different layer, on a different corpus slice — and reports the **same before
figures**: 264 reads and 10,740 bytes for
`ConditionalFormattingSamples.xlsx`, 26 reads and 2,478 bytes for
`sheet-names.xlsx`. This record's numbers are the pre-0573 state, and the two
agree with no adjustment.

That agreement is worth more than either number alone, because the two were
captured for different purposes: 0573 asks what the proof costs in reads on a
warm file, and this record asks what those reads cost a caller-supplied source.

## The read-ahead policies

`SourceReadPolicy::forward_start(N)` retains one window
(`crates/litchi-opc/src/source_backed/read_ahead.rs`, `ArchiveReadAhead::
forward_read`). On a miss it issues **one** physical read of
`min(source_length - offset, N)` bytes — never larger than the window, and
never sized by the caller's request — and returns a short count, so the ZIP
layer loops. A request that starts inside the retained window but runs past its
end is served short rather than refilling. Requests fall; bytes rise.

| fixture | exact | fs(4096) | fs(64K) | exact bytes | fs(4096) bytes | fs(64K) bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `comment.docx` | 15 | 4 | 4 | 2,114 | 5,809 | 6,639 |
| `endnotes.docx` | 13 | 5 | 3 | 3,571 | 13,375 | 14,640 |
| `testComment.docx` | 13 | 5 | 3 | 3,336 | 13,451 | 66,461 |
| `shape-glow-effect.pptx` | 19 | 5 | 3 | 4,254 | 13,531 | 14,934 |
| `shapes.pptx` | 49 | 14 | 3 | 10,830 | 52,880 | 69,264 |
| `shape-soft-edges.pptx` | 87 | 30 | 4 | 13,152 | 115,753 | 61,085 |
| `sheet-names.xlsx` | 40 | 6 | 3 | 5,537 | 17,282 | 10,323 |
| `universal-content.xlsx` | 58 | 8 | 4 | 4,398 | 20,873 | 9,501 |
| `ConditionalFormattingSamples.xlsx` | 354 | 77 | 49 | 35,252 | 308,776 | **2,953,536** |
| `SimpleNormal.xlsx` | 40 | 4 | 3 | 5,831 | 8,019 | 9,687 |
| `ExcelPivotTableSample.xlsx` | 86 | 15 | 3 | 11,706 | 52,251 | 21,509 |

On the 654,688-byte fixture the 64 KiB window asks the source for **2,953,536
bytes, 4.51 times the whole file**, to serve 49 requests. The 4 KiB window asks
for 308,776 bytes — less than half the file — to serve 77. **The larger window
is not uniformly better: 28 fewer requests, and 9.6 times the traffic.** On
this transport, where 64 KiB of transfer costs 0.625 ms against a 1 ms minimum
service, that traffic is nearly free; on a slower link it would not be.

The amplification is structural, not incidental. The 43 structural members of
the 132-member workbook are scattered across the archive, and each miss discards
the retained window and refills from the new offset: **37 fills for 43
members**, of which 34 are the full 65,536 bytes and three are shorter only
because they hit end of file. Only the proof sweep, which is sequential by
construction, collapses the way a window is supposed to — to **9 fills** over
the whole local-entry span.

Split by phase, that is where the two windows differ and where they do not:

| phase, 132-member fixture | exact | fs(4096) | fs(64K) |
| --- | ---: | ---: | ---: |
| open (scattered relationship parts) | 89 | 41 | **39** |
| cell read (sequential proof sweep, then the payload) | 265 | 36 | **10** |
| total | 354 | 77 | 49 |

Both windows collapse the sequential sweep well. Neither collapses the scattered
open below about 40 requests, because scattered access is exactly what a single
forward window cannot help.

## Gates

| Gate | Result |
| --- | --- |
| **determinism_gate** | **pass.** 44 (scenario × fixture × policy × route) arms, each compared across the three transports and across five repeats. **Zero divergent arms.** The request sequence is a pure function of the fixture, scenario and policy; neither the 1 ms service, the 100 MiB/s pacing nor the 64 KiB range cap changes one offset or one length. |
| **attribution_gate** | **pass, with a caveat stated below.** Every one of the requests across all 132 arms is classified. **Zero requests land in a gap; zero bytes are requested past end of file.** |
| **timing_gate** | **pass.** Delayed-transport wall clock is reported below as medians over five repeats with the host load average at capture. No warm-local latency is claimed. |
| **policy_gate** | **pass.** Read from source, listed below. |
| **control** | **pass.** For all eleven scenario/fixture pairs the two-step `SourceBackedPackage`-then-adopt route with the `exact` policy produces a **byte-identical ordered request sequence** to the native leaf constructor. The read-ahead arms are comparable. |

### The caveat on the attribution gate

The gate is "zero requests in gaps", and zero is the measured answer — but it is
weaker evidence than it sounds. The parsed layout of all eleven fixtures **tiles
every byte with no hole and no overlap**: end-of-central-directory, central
directory, and for each member a local header, variable header region, payload
and optional descriptor account for 100% of every file. A request therefore
*cannot* land in a gap on this corpus, whatever the library does. The gate would
only bite on an archive with prepended data, inter-member alignment padding or a
truncated tail, and this corpus has none. It is reported as passed and as
uninformative.

### policy_gate: read from source

Grepped over `crates/` at the pinned revision, **four public entry points in two
crates** accept a `SourceReadPolicy`:

| Entry point | File |
| --- | --- |
| `SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy` | `crates/litchi-opc/src/source_backed.rs:5580` |
| `SourceBackedPackage::..._and_source_read_policy_and_execution_context` | `crates/litchi-opc/src/source_backed.rs:5595` |
| `litchi_docx::source_backed::Package::from_read_at_with_limits_and_cache_limits_and_source_read_policy` | `crates/litchi-docx/src/source_backed.rs:543` |
| `litchi_docx::source_backed::Package::..._and_source_read_policy_and_execution_context` | `crates/litchi-docx/src/source_backed.rs:564` |

`litchi-pptx`, `litchi-xlsx`, `litchi-xlsb`, `litchi-ooxml-common` and the
`litchi` facade contain **zero occurrences of `SourceReadPolicy` in `src/`**. For
a presentation or a workbook the *only* way to attach a window is to build the
package with a policy and then call `from_source_backed_package`
(`crates/litchi-pptx/src/presentation/source.rs:954`,
`crates/litchi-xlsx/src/workbook/source.rs:431`), which takes no policy of its
own and inherits whatever the package was built with. That is the route this
record measures, and the control gate is what licenses comparing it to the
native constructors.

The pinned default is `SourceReadPolicy::exact()`, a private `window_bytes == 0`
(`crates/litchi-opc/src/source_backed/read_ahead.rs:29`, with
`Default` at `:32`). `forward_start`
accepts `1..=65_536` and rejects everything else.

## Delayed-transport wall clock

Medians over five repeats, 0493 transport, host load average **7.88** at the
start of capture and 7.43 at the end. **This is a model, not a device.**

| fixture | exact req | exact ms | fs(4096) req | fs(4096) ms | fs(64K) req | fs(64K) ms |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `comment.docx` | 15 | 16.07 | 4 | 4.44 | 4 | 4.45 |
| `endnotes.docx` | 13 | 14.73 | 5 | 6.24 | 3 | 4.26 |
| `testComment.docx` | 13 | 14.57 | 5 | 6.09 | 3 | 4.56 |
| `shape-glow-effect.pptx` | 19 | 20.23 | 5 | 5.47 | 3 | 3.44 |
| `shapes.pptx` | 49 | 52.25 | 14 | 15.26 | 3 | 4.16 |
| `shape-soft-edges.pptx` | 87 | 93.64 | 30 | 33.18 | 4 | 6.12 |
| `sheet-names.xlsx` | 40 | 42.45 | 6 | 6.50 | 3 | 3.39 |
| `universal-content.xlsx` | 58 | 61.53 | 8 | 8.68 | 4 | 4.50 |
| `ConditionalFormattingSamples.xlsx` | 354 | **375.14** | 77 | 82.85 | 49 | **81.24** |
| `SimpleNormal.xlsx` | 40 | 42.55 | 4 | 4.49 | 3 | 3.51 |
| `ExcelPivotTableSample.xlsx` | 86 | 91.93 | 15 | 16.90 | 3 | 4.37 |

Under the `exact` policy the cost is **1.060 to 1.133 ms per request** on every
fixture — the transport's 1 ms service plus sleep granularity, with the payload
transfer invisible because every request is small. The whole delayed-transport
result for `exact` is therefore the request count and nothing else.

The one place the byte amplification becomes visible is the 64 KiB window on the
132-member fixture: **1.658 ms per request**, against 1.060 for `exact`, because
a 64 KiB fill costs 0.625 ms of transfer on top of the 1 ms service. That is why
`forward_start(65536)` (49 requests, 81.24 ms) and `forward_start(4096)` (77
requests, 82.85 ms) land **within 1.9% of each other** despite a 28-request
difference. The 64 KiB window is the faster of the two, by 1.61 ms — but this
record's own capture-to-capture spread reaches 1.33%, so 1.9% is barely outside
its own noise. The defensible statement is that on this scenario and this
transport the two windows are **indistinguishable**: what the larger one saves
in requests it returns in traffic. It is not a measurement that one is better.

Two independent captures of the whole matrix, minutes apart at load averages
7.88 and 7.21, agree **within 1.33% at every one of the 33 medians**, with
request counts identical arm for arm
([`timing-repeatability.json`](results/change-0572/timing-repeatability.json)).
The transport is sleep-dominated, which is why an eight-deep load average on a
32-core host does not move it. That is a statement about this model's
insensitivity to host load, **not** evidence that the library's own CPU time is
insensitive to it — that was never measured.

## Predictions: six confirmed, five wrong, one half

The plan's `predicted_effect` entries, scored against measurement.

### Confirmed

1. **`transport_independence`** — "the ordered request sequence of every scenario
   is identical between the zero-delay and the 1 ms plus 100 MiB/s transports and
   across at least five repeats". **Exactly right.** 44 arms, zero divergence,
   and the extra range-cap arm found none either.

2. **`docx_text_exact`** — "two more requests for `word/document.xml` (three when
   it carries a descriptor); no strict-layout proof". **Exactly right**, on all
   three DOCX fixtures.

3. **`pptx_slide_exact`** — "two to four more requests for the selected slide part
   (plus descriptors)". **Right**; measured two, or three with a descriptor.

4. **`xlsx_cell_exact`, the proof size** — "two requests per member, three per
   descriptor-bearing member … 26 on `sheet-names.xlsx`, 264 on
   `ConditionalFormattingSamples.xlsx`, 39 on `universal-content.xlsx`", "issued
   before the first worksheet payload byte", "more than 85 percent of the cell
   phase". **All exactly right**: 26, 264, 39, contiguous, ascending in layout
   order, ending immediately before the first worksheet payload read, and 89.7%,
   97.5% and 99.6% of the cell phase respectively. The "more than 85 percent"
   holds on every fixture the plan named; it is 82.8% on `SimpleNormal.xlsx`,
   which the plan did not name.

5. **`xlsx_cell_exact`, what follows the proof** — "then worksheet,
   shared-strings and styles payload reads". **Right, on the fixture that
   exercises all three.** On `ExcelPivotTableSample.xlsx` the seven requests
   after the proof are, in order: the worksheet payload, the worksheet local
   header and payload *again*, the shared-strings local header and payload, and
   the **styles** local header and payload. The list is read on demand, not
   unconditionally: `SimpleNormal.xlsx` reads worksheet and shared strings but
   not styles (five requests), and `sheet-names.xlsx`, whose A1 is absent, reads
   only the worksheet (three). No fixture reads a part the prediction did not
   name.

   **An earlier draft of this record scored this prediction as wrong**, on the
   strength of `SimpleNormal.xlsx` alone, and stated that no styles payload is
   ever read. That was contradicted by this record's own retained bytes, and was
   caught in review before publication. It is recorded rather than quietly
   corrected.

6. **`read_ahead_64k`, the small-fixture cost** — "at the cost of window-sized
   fills that read most or all of a small fixture". **Right, on all ten.** Every
   fixture below 100 KB is read at **101% to 135% of its own size** under
   `forward_start(65536)`: 1.01x, 1.02x, 1.08x, 1.09x, 1.09x, 1.10x, 1.11x,
   1.22x, 1.30x, 1.35x.

### Wrong

1. **`open_exact`, the structural-member count** — predicted "three structural
   members for DOCX and XLSX, twenty-one for the six-slide `shapes.pptx`".
   Measured **3 or 4 for DOCX, 4 to 43 for XLSX, and 22 for `shapes.pptx`**. The
   replacement rule above — content types, every `.rels` part, and the main part
   — predicts all eleven fixtures exactly. On the 132-member workbook the count
   is 43, fourteen times the prediction and 89 of that scenario's 354 requests.
   The per-member *shape* the same prediction gives is correct; only the count
   is wrong, and it is wrong because the plan treated a package property as a
   format constant.

2. **`xlsx_cell_exact`, the scenario share** — predicted the proof is "roughly 95
   percent of the whole scenario on the 132-member fixture". Measured **74.6%**.
   The prediction omitted the 89-request open the same plan mis-sized.

3. **`xlsx_cell_exact`, the total** — predicted "about 264 of roughly 280 ms at
   1 ms per request". The scenario is **354 requests** and its measured median is
   **375.14 ms**, not about 280, for the same reason.

4. **`read_ahead_4k`, the proof floor** — predicted "the 132-member proof is
   predicted to remain above 100 requests" under `forward_start(4096)`. Measured
   **36 requests for the whole cell phase**, proof included — under half the
   predicted floor.

5. **`read_ahead_64k`, the scenario total** — predicted "the whole 132-member
   XLSX scenario to below 25 requests". Measured **49**. The sweep half of that
   prediction is nearly exact — "about one fill per 64 KiB of local-entry span,
   about 11" against a measured 9 — but the open phase costs 39 requests on its
   own, so the scenario cannot reach 25.

### Half right: the split mechanism

`read_ahead_4k` also says "a request larger than the window is served in
window-sized fills, so a member payload larger than 4 KiB costs more requests
than exact". That is **two claims, and they score differently.**

The **mechanism is confirmed.** Exactly two reads in the whole corpus exceed the
4 KiB window, and both are split:

| read | `exact` | `forward_start(4096)` |
| --- | ---: | ---: |
| 9,724-byte central directory, `ConditionalFormattingSamples.xlsx` | 1 request | **3 fills** |
| 5,117-byte central directory, `shape-soft-edges.pptx` | 1 request | **2 fills** |

The **consequence is not observable here**, because the reads it names do not
exist in this corpus: **no member payload read anywhere in these eleven
fixtures exceeds 4,096 bytes.** The largest is 1,251 bytes, on the 655 KB
workbook. These scenarios only ever touch content types, relationship parts and
one target part, all small XML; no media member is ever read. So no arm has
`forward_start(4096)` costing more requests than `exact` — the two split
central directories cost three extra requests between them, against savings of
hundreds.

A corpus with media-bearing parts on the read path would test the consequence.
This one cannot, and the record does not pretend otherwise.

## A correction to change 0567

Change 0567 states "Every structural member is read exactly once per open —
three for DOCX and XLSX, twenty-one for a six-slide PPTX". The *once per open*
claim holds. The counts do not: they are the counts of a particular fixture
triple, not of the formats, and even for `shapes.pptx` the measured number here
is 22 rather than 21. Change 0567's conclusion — one archive index per open — is
untouched by this.

## Three things found that the plan did not ask about

**A DOCX limit that stops full-text extraction.** `write_text_to` refuses
`drawing.docx` (22 members, the largest DOCX member count in the corpus) and
`table-alignment.docx` with
`InvalidFormat("semantic DOCX XML exceeds 4096 namespace bindings")`. Those two
fixtures were therefore dropped from the corpus and the largest DOCX measured is
17 members. This is recorded, not diagnosed.

**An XLSX streaming-scan refusal that reaches the caller.** `SourceWorksheet::
cell("A1")` returns `Invalid("worksheet mergeCells appears before sheetData")`
on `universal-content.xlsx` and `ConditionalFormattingSamples.xlsx` — **two of
the three fixtures the plan itself named.** Every `.xlsx` fixture directly under
`test-data/ooxml/xlsx` was then driven through the same scenario to size it
([`xlsx-cell-survey.json`](results/change-0572/xlsx-cell-survey.json)):

| outcome of `cell("A1")` on sheet 0 | fixtures |
| --- | ---: |
| completed | 82 |
| refused, `mergeCells appears before sheetData` | **10** |
| refused, `shared-string count exceeds 2147483647` (a deliberately malformed fixture) | 1 |
| **total surveyed** | **93** |

The strict-layout proof still runs in full before the refusal, so the request
counts above are sound; but the scenario does not return a cell on those ten.
`SimpleNormal.xlsx` and `ExcelPivotTableSample.xlsx` were added to the corpus
beyond the plan's named set precisely so the proof is also observed on a path
that completes with a stored value, and it is: 24 and 54 proof requests, same
shape. Whether the streaming scanner should fall back to the eager store here
rather than refuse is a question for someone else; this record reports only that
it affects roughly one XLSX fixture in nine, and that the plan drew two of its
three from that group by chance.

**The selected worksheet member is read twice.** After the proof, the cell read
issues a payload read of the worksheet, then reads the *same member's* local
header and payload again. Two of the 29 requests on `sheet-names.xlsx` and two
of the 29 on `SimpleNormal.xlsx` are that repeat. It is small, it is consistent,
and it is not explained here.

## Limitations

One host, one toolchain, warm in-memory sources, eleven fixtures, five repeats.

The source is owned bytes, not a file. The counting wrapper holds each fixture in
memory, so the control arm measures the library and the model and never the page
cache. No filesystem, cold-cache, physical-device or cross-platform result is
claimed, and no allocation, peak-RSS, instruction-count or syscall measurement
was taken. The request counts are logical `ReadAt` calls, which under the
`exact` policy are one-to-one with what `soapberry-zip` asks for, and under a
`forward_start` policy are the physical fills the window issues.

No timing is claimed beyond the model. The delayed medians are sleep-driven
arithmetic over a deterministic request count. The host carried unrelated load
throughout the session; the load average at capture is recorded in the result
JSON. **Warm-local latency is explicitly not claimed**, and no ABBA comparison
was run because this record proposes no candidate to compare.

The proof detector is mechanical and has one boundary artifact. It reports the
longest run of consecutive requests that either begin exactly at a member's
local-header offset or read only local-record framing. On
`universal-content.xlsx` that run is 40 requests because it swallows the
preceding structural member's descriptor read; the proof itself is 39, three per
member for thirteen members, and this record uses 39 throughout.

The gap check is uninformative on this corpus, for the reason given under the
attribution gate. Nothing under `crates/` was touched: the probe is a throwaway
binary with path dependencies on a pinned extract of the tree, retained with the
evidence.

## Disposition

Attribution only. Nothing is proposed, nothing is landed, and no follow-up is
opened by this record. Its purpose is to put a price on two shapes that change
[0561](0561-opc-repeated-positional-reads.md) and
[0567](0567-ooxml-single-index-per-open.md) had described but not measured on a
range source, so that later work can be argued against numbers.

Two of those numbers are already being acted on elsewhere: change
[0573](0573-zip-single-local-header-read.md) halves the proof, and change
[0575](0575-zip-lazy-strict-layout-design.md) is the design record for making it
lazy. The figure this record contributes that neither of those addresses is the
**open**: 89 of the 132-member workbook's 354 requests, scaling with the
relationship-part count, and the one phase neither read-ahead window collapses
below about 40 requests.

Evidence packet, including the probe, the analysis scripts and the full ordered
request sequences: [`results/change-0572/README.md`](results/change-0572/README.md).
