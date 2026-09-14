# 0569: the detect-then-open saving, priced, and a correction to 0567

Status: attribution only. No production change and `performance_claim: none`.
It corrects the recommendation change
[0567](0567-ooxml-single-index-per-open.md) made, corrects one of its figures,
and records a second defect found while checking it.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What 0567 recommended, and why it is wrong

Change 0567 found that `detect_file_format(path)` followed by a facade open
indexes the ZIP container twice, and proposed two fixes. It said the cheaper one
should be priced first:

> A zero-code alternative worth pricing first. `Document::open`,
> `Presentation::open` and `open_workbook` already detect internally and already
> classify a non-matching OOXML family, so documentation telling callers not to
> pre-call `detect_file_format` captures the same saving with no production
> change and no risk.

Three claims in that sentence fail.

**"All three facade opens already detect internally."** `litchi::sheet::open_workbook`
does no detection at all. It calls `SourceBackedWorkbook::from_path_with_limits`
directly and therefore fails on `.xls`, `.ods` and `.xlsb` — three formats that
`Workbook::open` handles. The correct target for that advice is `Workbook::open`,
not `open_workbook`.

**"Already classify a non-matching OOXML family."** The classification is
computed and then discarded. Ten call sites read `let _ = format;`. The error
type has nowhere to put it: `litchi-core` carries only an `InvalidFormat(String)`
and a unit `NotOfficeFile`.

**"No risk."** A caller who pre-detects in order to choose which facade type to
build loses their only supported mechanism. There is no unified opener, and
`Document`, `Presentation` and `Workbook` share no trait.

## What the facades actually report for mismatched input

No facade error names the real format. The two wrong-family messages are
byte-identical constant strings across all six wrong formats tested.

| Input | `detect_file_format` | `Document::open` | `Presentation::open` | `Workbook::open` |
| --- | --- | --- | --- | --- |
| DOCX | `Docx` | opens | wrong-family constant | **`NotOfficeFile`** |
| PPTX | `Pptx` | wrong-family constant | opens | **`NotOfficeFile`** |
| XLSX | `Xlsx` | wrong-family constant | wrong-family constant | opens |
| legacy `.doc` | `Doc` | opens | wrong-family constant | **`NotOfficeFile`** |
| legacy `.ppt` | `Ppt` | wrong-family constant | opens | **`NotOfficeFile`** |
| ODT | `Odt` | opens | wrong-family constant | **`NotOfficeFile`** |
| PNG or CSV | none | `NotOfficeFile` | `NotOfficeFile` | `NotOfficeFile` |

`Workbook::open` returns the identical unit `NotOfficeFile` for a Word document,
a legacy Word document, an ODT, an ODP, a PNG and a CSV. It cannot distinguish
"this is a Word document" from "this is not an Office file", where
`detect_file_format` separates them cleanly. That is a strict information
regression for any caller told to stop pre-detecting.

`Document::open` and `Presentation::open` do split coarsely, reporting
`NotOfficeFile` for garbage and a wrong-family error otherwise, so they give a
two-way signal but never the identity. `open_workbook` leaks raw OPC and ZIP
text and reports a legacy `.xls` with the same message as a PNG.

## What the saving is worth, measured

Paired A/B interleaved in one loop, warm cache, pinned, two independent runs at
2,000 and 3,000 iterations on different cores agreeing within 0.6%.

| Pattern | two-call | single-call | saving | share of two-call |
| --- | ---: | ---: | ---: | ---: |
| DOCX `Document::open` | 87.45 us | 44.91 us | **42.2 us** | 48.6% |
| PPTX `Presentation::open` | 474.51 us | 254.88 us | **219.7 us** | 46.3% |
| XLSX `Workbook::open` | 128.65 us | 79.00 us | **49.3 us** | 38.6% |
| XLSX `open_workbook` | 115.97 us | 67.03 us | **48.9 us** | 42.2% |

Segmenting `strace` per `openat` of the fixture — rather than per descriptor,
which is the error change 0567 corrected in change 0561 — confirms the
mechanism: the single-call pattern opens the package once, the two-call pattern
twice, at 2.2 to 2.5 times the package reads.

### A correction to 0567's own figure

This reproduces change 0567 within 2.5% on every isolated figure, but its
projected PPTX saving of 202.4 microseconds is **low**. The measured paired
delta is **219.7 microseconds**, because the two-call pattern also pays a second
`openat` and loses warmed state. Change 0567's figure was the arithmetic
difference of two separately measured medians; this one is a paired measurement
of the two patterns themselves. The DOCX and XLSX projections were close.

Change 0567 also attributed 65.51 microseconds to "the XLSX open". That was
`open_workbook`. `Workbook::open` costs 78.4 microseconds, about 12 more,
because it also projects core properties.

## Pre-detecting is unavoidable for the largest affected population

The obvious workaround — try each facade in turn — is measured and worse:

| Fixture | detect then open | try each facade | difference |
| --- | ---: | ---: | ---: |
| DOCX, tried first | 87.13 us | 44.63 us | −42.50 us |
| PPTX, tried second | 479.94 us | 472.35 us | −7.59 us |
| XLSX, tried third | 126.26 us | 175.81 us | **+49.55 us** |

A failed `Document::open` on a presentation costs about 216 microseconds,
because it builds a full index before failing. The cascade is a lottery on file
ordering, and it also destroys the coarse `NotOfficeFile` signal.

`detect_format_smart` is already a dead end for this: it is public, but
`Document::from_detected` and `Presentation::from_detected` are both private, and
it takes owned bytes so it forfeits source-backed path opens.

## The advice has far less reach than the saving suggests

An exhaustive sweep found that **no Rust doc comment anywhere tells callers to
detect before opening**, and the root readme never mentions `detect_file_format`
at all. The only place in the repository that teaches the pattern is
`crates/litchi/examples/to_markdown.rs`, at lines 39 and 289 to 300 — which is
simultaneously the canonical anti-pattern and the unavoidable dispatch case.
**Under the proposed advice that example has no correct rewrite.** The detection
surface is also exported as public Python API through `pyo3-litchi`.

## Disposition

| Recommendation | Status |
| --- | --- |
| "Callers should not pre-call `detect_file_format`" as a blanket statement | **Rejected.** Unsafe for any caller dispatching across the three facades, any caller needing the format identity, and the Python surface. |
| "If you already know the expected family, call the facade opener directly" | **Safe**, and captures the full saving. Extension-based routing before the open is free. |
| Same advice naming `open_workbook` | **Unsafe.** It is XLSX-only and undetected; the correct target is `Workbook::open`. |
| The opaque prepared-source entry point | **The only fix that recovers the time for the unavoidable case**, and it now has a measured ceiling rather than an estimated one. |
| Surfacing the discarded classification in a typed error | **Correct and independently worth doing**, but it is not a performance fix: reaching the error still costs a failed open. |

## Limitations

Warm cache, one host, one fixture per format, one feature set, one cascade
ordering. No cold-cache, latency-bearing-source, allocation or throughput result
is claimed. Host load average was 21 to 25 from concurrent builds, mitigated by
pinning and by A/B interleaving in one loop, so the absolute medians may run a
few percent high while the paired deltas are drift-cancelled by construction.
Every error value and every timing above is measured; the call-site sweep and
the ADR reasoning are read from source. Nothing here is estimated.
