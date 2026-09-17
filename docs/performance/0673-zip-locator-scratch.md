# 0673: the managed ZIP index defers its 64 KiB locator scratch

Status: retained, implemented in `soapberry-zip`. `performance_claim: none`.
This record completes change 0651's row 18, which asked for a measurement of
the 64 KiB locator scratch that change 0632 left on the managed index path.

The managed `IndexedArchive::from_reader*` path now probes the fixed terminal
EOCD through a 46-byte stack buffer. A valid zero-comment EOCD returns through
the existing central-directory prefill path without first allocating the
locator's 64 KiB search window. If that probe misses, the locator fallibly
allocates the same 64 KiB window before its backwards search and uses it for
the unchanged search and parse path. The public `ZipLocator` entry points keep
their default behavior and buffer contract.

The change is deliberately bounded. It does not infer a new ZIP shape, skip a
validation, or alter the central-directory prefill. The fallback retains the
historical search-window size, and the existing fixed-record, comment, ZIP64,
prefix-offset, short-read, limit, and typed refusal paths remain in place.

## Authority and scope

Change 0651 row 18 names the remaining locator scratch and prices a two-stage
locate as an extra request on the one corpus container that misses the
`len - 22` fast path. Change 0632's packet records the prior prefill design
and leaves this two-stage choice unimplemented. Change 0652's correctness and
safety standing decision applies: the common benign path may defer work needed
only by exceptional inputs, while those inputs retain their typed refusal
before a partial result.

The production diff is limited to `crates/soapberry-zip/src/locator.rs` and
`office.rs`. The focused test adds an assertion that the exact path keeps the
22-byte EOCD probe followed by the central-directory window, while a comment
forces the 64 KiB fallback window. No public API, error kind, archive limit,
ZIP record, or output member is changed.

## Measurement

The packet reuses change 0632's allocation/request probe, builds both legs in
release mode from the shared before checkout and this worktree, and pins the
probe to CPU 8. The three fixture open costs are:

| fixture | open allocations | allocated bytes | requests / returned bytes | source versions |
| --- | ---: | ---: | ---: | ---: |
| `xlsx-132` | 5,338 → 5,337 | 2,256,370 → 2,190,833 | 9 / 26,020 → 9 / 26,020 | 4 → 4 |
| `pptx-shapes` | 1,995 → 1,994 | 541,256 → 475,719 | 5 / 11,145 → 5 / 11,145 | 4 → 4 |
| `docx-comment` | 403 → 402 | 365,976 → 300,439 | 5 / 1,644 → 5 / 1,644 | 4 → 4 |

The allocation result is one fewer allocation and 65,537 fewer allocated bytes
per fixture. It is deterministic support for removing the managed locator
scratch; no latency or speed claim is registered.

The open oracle covers the 533-container list retained by change 0632. Its
before and after reports each contain 18,494 lines and are byte-identical,
including request, byte, and source-version costs. The oracle therefore
records no changed verdict, error identity, relationship, part, non-part
member, decoded payload, or request grammar. The per-fixture test forces the
comment-bearing missed-probe path that does not occur in the corpus's open
costs.

## Correctness and limits

The locator's fallback uses `try_reserve_exact` and the existing typed
`ErrorKind::Allocation` resource. A failed or rejected fixed probe still falls
through to the established backwards search, and the returned archive still
passes through the same EOCD, ZIP64, central-record, limit, and directory scan
validation. The only changed common-path state is the storage location of the
46-byte probe.

The focused `soapberry-zip` suite covers exact and comment-bearing archives,
false EOCD signatures in comments, prefixed archives, ZIP64 fixtures, capped
short reads, undersized directories, oversized variable fields with their
`BufferTooSmall` refusal, and empty archives. The full `soapberry-zip` and
`litchi-opc` suites pass as well.

No cold-cache, physical-device, remote-latency, peak-RSS, instruction-count,
concurrency, cross-platform, or allocation-failure-injection result is
claimed. The 64 KiB fallback remains a bounded cost for a missed probe.

Reproduction material is retained in
[`results/change-0673/`](results/change-0673/), including the paired probe,
the oracle runner and summaries, counts, provenance, and the four coordinator
log sections.
