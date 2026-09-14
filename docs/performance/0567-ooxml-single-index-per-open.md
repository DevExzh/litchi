# 0567: one OOXML open builds one archive index, and two corrections

Status: attribution only. No production change and `performance_claim: none`.
It disproves a standing hypothesis, corrects two statements in changes
[0561](0561-opc-repeated-positional-reads.md) and
[0562](0562-zip-descriptor-read-once.md), and locates where repeated indexing
really happens.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## The question

`docs/GOAL.md` hypothesis 12 reads "Detection may repeat container indexing
unless an opaque prepared source is reused." Change 0561 appeared to confirm it:
segmenting a traced perf-harness child's positional reads at each 22-byte
end-of-central-directory record found **ten** archive constructions, and since
the child performs roughly three package opens, the record inferred that each
library-level open builds the ZIP index about three times.

## The answer: one

Every library-level open measured builds **exactly one**
`soapberry_zip::office::IndexedArchive`. Leaf constructors and the `litchi`
facade alike, for DOCX, PPTX and XLSX, under two feature sets, verified by two
independent methods — 26 debugger runs breaking on the constructors and printing
backtraces, and 26 `strace` captures counting end-of-central-directory reads on
the descriptor that actually holds the package. All 52 report one.

| Scenario | Library call | Constructions |
| --- | --- | ---: |
| DOCX leaf open, and with text or one paragraph | `litchi_docx::source_backed::Package::open` | **1** |
| DOCX facade open, and with text | `litchi::Document::open` | **1** |
| PPTX leaf and facade open, and with one slide | `SourceBackedPresentation::open`, `litchi::Presentation::open` | **1** |
| XLSX leaf and facade open, and with one cell | `SourceBackedWorkbook::open`, `litchi::sheet::open_workbook` | **1** |
| DOCX from owned bytes | `litchi::Document::from_bytes` | **1** |

Inside one open, hypothesis 12 is **already satisfied**. The facade builds the
`SourceBackedPackage` once during detection and hands that same prepared package
to the format owner:
`litchi::detection_smart::detected::detect_docx_source_with_limits` constructs
it, `try_detect_ooxml_format_from_source_backed_package` classifies it, and
`litchi_docx::source_backed::Package::from_source_backed_package` adopts it
without reparsing. Every structural member is read exactly once per open — three
for DOCX and XLSX, twenty-one for a six-slide PPTX — so the part-cache bypass in
`read_structural_member` costs nothing *within* an open.

Each construction reads the end-of-central-directory record once, then the
central directory in two reads: a 46-byte first-record probe, then the whole
directory. That is 751 bytes for the DOCX fixture, 3,730 for the PPTX and 900
for the XLSX.

One archive construction is 3.9% to 6.8% of one library-level open. With exactly
one per open, **the wall-clock share of redundant construction inside one open is
zero.** The facade costs 3.5 microseconds more than the leaf on DOCX and 4.1 on
PPTX, which is a `FileSource::open` plus two four-byte magic probes and an ODF
`mimetype` name sniff, not a second index.

## Correction to change 0561: ten constructions is two per process

Change 0561's ten constructions per traced child are real, but they are not ten
per process. Re-reading its own retained traces with process and descriptor
attribution shows the ten end-of-central-directory reads spread across **five
processes, two each**, in all four captures.

| Capture | Constructions per process |
| --- | ---: |
| `docx_file_source_open` | 2 |
| `docx_file_source_full_text` | 2 |
| `pptx_file_source_open` | 2 |
| `pptx_file_source_selected_slide` | 2 |

Four are the per-sample children and the fifth is the parent. Two per process is
exactly the lifecycle change 0561 itself describes: the harness opens the package
once untimed to prepare and once in the post-timer oracle. The counts were never
wrong; the inference drawn from them was. Each library-level open builds the
index **once**, not about three times.

Plain `(pid, fd)` segmentation is unsafe on these traces, which is how the
original reading arose: the dynamic loader uses descriptor 3 before `main`, and
the package file later reuses descriptor 3. Attribution has to follow `openat`
and `close` and count only reads on the descriptor holding the package.

## Correction to change 0562: the 840-byte read is the dynamic loader

Change 0562 recorded that its DOCX captures retain 8.1% to 8.5% immediate
duplicate reads after the change, described them as "one 840-byte read at offset
64 repeated two to four times consecutively", and attributed them to "a
structural member being re-read", listing it as an open follow-up.

That attribution is wrong. With stack-resolving tracing the read is:

```
pread64(3, "\6\0\0\0\4\0\0\0@\0\0\0…", 840, 64) = 840
 > ld-linux-x86-64.so.2(open_verify.constprop.0+0x1fb)
 > ld-linux-x86-64.so.2(_dl_map_new_object+0x5f2)
 > ld-linux-x86-64.so.2(dl_main+0x1c10)
```

The preceding `openat` is `/usr/lib/x86_64-linux-gnu/libc.so.6`. It is the
dynamic loader reading libc's ELF program-header table before `main` runs, on a
transient descriptor 3 that the package file later reuses. It is not
`[Content_Types].xml`, not a relationship part, and not attributable to this
library at all. Two per process, 24 across the eleven processes in each retained
trace.

The follow-up change 0562 opened on that basis is therefore closed, not
deferred.

## Where indexing really is repeated: across two calls

The second constructor site,
`probe_package_catalog_from_reader_with_limits` at
`crates/litchi-opc/src/pkgreader.rs:307`, is reached only from the
reader-and-bytes detection façade and never from a path-based open. So
`detect_file_format(path)` followed by an open **does** index twice, re-reads the
central directory twice, and re-reads every structural member twice — twenty-one
of them for a six-slide presentation.

| Fixture | detect | open | total | detection's share |
| --- | ---: | ---: | ---: | ---: |
| `Hyperlink.docx` | 42.11 us | 45.13 us | 87.85 us | 47.9% |
| `shapes.pptx` | 202.41 us | 256.34 us | 458.75 us | 44.1% |
| `sheet-names.xlsx` | 46.98 us | 65.51 us | 112.49 us | 41.8% |

Detection on a `Read + Seek` reader has no way to publish the prepared package to
a subsequent open, so it indexes into a throwaway borrowed reader and drops it.
The remaining constructor sites are a validation-only entry point that none of
these paths reach, and a test fixture.

## A separate cost found on the XLSX path

`litchi_xlsx`'s streaming cell read reaches
`IndexedArchive::with_verified_entry_reader`, which builds a strict layout proof
that DOCX and PPTX never trigger. The proof walks every entry and issues **two
positional reads per member**, at a very consistent 319 to 345 nanoseconds per
member across five fixtures. It is memoized per archive, so a second cell read
costs nothing.

| Fixture | Members | Isolated proof cost |
| --- | ---: | ---: |
| `sheet-names.xlsx` | 13 | 4.45 us |
| `Tables.xlsx` | 46 | 15.26 us |
| `ConditionalFormattingSamples.xlsx` | 132 | 42.94 us |

On the smallest fixture that is 6.2% of the open-to-first-cell interval. The
proof re-reads local-header framing the central directory already describes,
which is the class change 0561 identified, but it establishes a whole-archive
non-overlap property rather than a per-entry one, so it cannot be made per-entry
without weakening it. This is recorded, not addressed: it is a correctness
-sensitive security check and a 4.45 microsecond warm-cache saving does not
justify touching it. Change 0561's own governing conclusion applies — justify
such work against a latency-bearing source, not a warm-cache count.

## What a fix would look like

ADR 0010 anticipates this case directly, requiring that optimization here belong
to "a focused, opaque detection plan or neutral container capability" with raw
archive types kept out of the facade, and its 2026-08-08 amendment records the
accepted shape as one classified immutable package snapshot consumed by exactly
one format adapter. ADR 0011 keeps `litchi-opc` the only owner translating OPC to
the ZIP implementation.

The OOXML analogue is already built internally; what is missing is a public entry
point that returns the prepared package instead of discarding it. The smallest
compliant change adds an opaque newtype in `litchi-opc` wrapping the existing
`SourceBackedPackage` and its validated main-part content type, a
`detect_and_prepare` function beside `detect_file_format` implemented through the
existing path-based detection, and `from_prepared` constructors on the three
facade types forwarding to the already-public adopters. `soapberry-zip` is
unchanged.

The projected saving is 42.1 microseconds per DOCX, 202.4 per PPTX and 47.0 per
XLSX for any caller that detects then opens. That figure is **estimated**, being
the difference of two measured numbers; no such API exists to benchmark.

A zero-code alternative should be priced first. All three facade opens already
detect internally and already classify a non-matching OOXML family, so
documentation telling callers not to pre-call `detect_file_format` captures the
same saving with no production change and no risk.

> **Corrected by change [0569](0569-ooxml-detect-then-open-priced.md).** Three
> claims in the paragraph above are wrong. `litchi::sheet::open_workbook`
> performs no detection at all and fails on three formats `Workbook::open`
> handles. The classification is computed and then *discarded* at ten call
> sites, so no facade error names the detected format, and `Workbook::open`
> reports an identical unit `NotOfficeFile` for a Word document, a legacy
> document, an ODF file, a PNG and a CSV. And the advice is not riskless: there
> is no unified opener, so a caller dispatching across the three facade types
> must pre-detect, and this repository's own flagship example would have no
> correct rewrite under it. Change 0569 also measures the PPTX saving at 219.7
> microseconds against the 202.4 projected below, and identifies the 65.51
> microsecond XLSX open below as `open_workbook` rather than `Workbook::open`.

## Limitations

Warm cache, one host, single-fixture medians per format. No cold-cache,
latency-bearing-source, allocation, throughput or multi-fixture-corpus result is
claimed. A `perf record` run produced a flat profile that could not resolve
inlined frames; the wall-clock figures are paired microbenchmark medians over
1,000 to 3,000 iterations, not profile attributions. ODF appears only as the
`mimetype` name sniff on the OOXML facade path and is otherwise unmeasured;
iWork is excluded.
