# 0510 ADR review

Reviewed 2026-09-11 against [`docs/adr/README.md`](../../../adr/README.md) and
all 29 numbered records. The supplied [ADR hash manifest](adr-manifest.json)
has 30 entries (the records plus the README); recomputing SHA-256 over the
current bytes found zero mismatches. Its recorded revision is
`869bbc3ed470453b8324893180360a2d69e7c315`.

The candidate is a private change to
[`normalized_xml10_decoded_len`](../../../../crates/litchi-odt/src/elements/text.rs)
and its focused tests in [`sequential_text.rs`](../../../../crates/litchi-odt/tests/sequential_text.rs).
It keeps `str::from_utf8` validation first, uses one `memchr` search for the
first carriage return, returns the exact raw length when no carriage return is
present, and otherwise enters the existing scalar CRLF walk at that first
carriage return. No public signature, writer call, event traversal, or source
mutation is changed. The diff adds no unsafe code, manifest entry, dependency,
or lockfile change.

The source audit finds these invariants preserved:

1. UTF-8 validation and its `InvalidFormat` context remain before the fast
   path. CR, CRLF, repeated CRLF, and lone CR handling use the same scalar
   logic after the first CR; only a no-CR prefix can return early.
2. The decoded length still feeds the existing `SinkTextBudget::charge` and
   `append_sink_precharged` checks. `MAX_TEXT_BYTES`, depth/block/space limits,
   fallible reservations, and `SequentialTextWriter` output/object limits are
   unchanged.
3. The helper has no access to `PendingSinkBlocks` or the writer. Nested block
   start order, pending frontiers, output progress, and writer boundaries are
   therefore unchanged. Existing source-backed output remains read-only and
   line-ending normalization remains the established semantic behavior.
4. The added differential and integration cases cover invalid UTF-8, CR/CRLF
   edge patterns, Unicode, exact normalized output, source byte retention,
   exact output limits, and sink progress. They were inspected but not run by
   this documentation audit.

## Admission and guardrails

The [admission record](admission.json) admits a *trial* first-CR hybrid and
explicitly makes final retention conditional on matched export and newline
guardrails. The frozen control profile is the recorded binary with
298,831,445 Callgrind instruction references; the helper accounts for
28,350,000 exclusive references (9.49%) in that profile. The isolated
[helper guardrail](helper-guardrail/README.md) rejects unconditional
`memchr_iter` (up to +672.3% p50) and `memmem_iter` (up to +754.3% p50). The
selected hybrid has remaining helper flags of +33.3% on tiny dense CRLF and
+8.4% on large sparse CRLF, so those flags remain part of the end-to-end
decision.

The pending public guardrail covers nine newline exports, freezes the control
binary and exact preflights before editing, and keeps CountSink timing distinct
from the hashing baseline. At this review point the retained artifacts contain
the frozen [before preflight](export-guardrail/before-preflight.csv) and no
matched after result. This review consequently records no final latency,
throughput, allocation, cache, branch, or RSS improvement and does not promote
the trial to an unconditional performance admission. The available hardware
probe only establishes that the requested event probes can run; it is not an
after-export comparison.

## Applicability matrix

| ADR | Applicability and disposition |
|---|---|
| 0001 | **Direct.** Correctness, safety, typed failures, and strict public layers remain primary. The helper is private and adds no raw/public type, panic path, or unsafe code. |
| 0002 | **Direct boundary.** The edit stays in `litchi-odt`; it adds no peer-family, umbrella, or dependency edge and changes no ownership direction. |
| 0003 | **Applicable.** The helper reads immutable event bytes and changes no snapshot, edit, patch, lineage, or source state. |
| 0004 | **Direct API constraint.** Semantic text and the existing sequential-writer API are unchanged; no public signature or facade type is added. Plain text remains inert. |
| 0005 | **Direct.** The optimization is measured under the existing sequential sink and retains typed precharge, budget, output-limit, and progress boundaries. The helper profile is diagnostic; final performance admission remains gated by the matched nine-case export guardrail. |
| 0006 | **Direct.** UTF-8 validation, XML 1.0 normalization, decoded-byte accounting, typed error context, and source-preserving read behavior remain. The fast path does not rewrite markup or activate inert content. |
| 0007 | **Direct ODF model constraint.** Visible paragraph/heading text, controls, tracked-change exclusion, note/ruby suppression, nested start order, and paragraph output semantics are unchanged. |
| 0008 | **Direct evidence gate.** The helper differential and frozen public preflight establish the candidate’s evidence boundary. Final retention still requires matched output, progress, malformed-input, failure, and newline-performance checks; no broader support claim follows. |
| 0009 | **No direct change.** ODF detection ownership and its fuzz boundary are untouched. |
| 0010 | **No direct change.** Facade and archive ownership are untouched. |
| 0011 | **No direct change.** OOXML physical package ownership is untouched. |
| 0012 | **No direct change.** BIFF8 formula reference types and encoding are untouched. |
| 0013 | **No direct change.** PPTX notes ownership and deletion are untouched. |
| 0014 | **No direct change.** The amended core-properties reader ownership record is unaffected. |
| 0015 | **No direct change.** Lossless OOXML core-properties CRUD is unaffected. |
| 0016 | **No direct change.** BIFF8 writer-location types are untouched. |
| 0017 | **No direct change.** OOXML producer-template ownership is untouched. |
| 0018 | **No direct change.** XLSX calculation-chain ownership is untouched. |
| 0019 | **No direct change.** DOCX web-settings ownership is untouched. |
| 0020 | **No direct change.** PPTX table-style ownership is untouched. |
| 0021 | **No direct change.** DOCX glossary/building-block ownership is untouched. |
| 0022 | **No direct change.** PPTX embedded-font ownership is untouched. |
| 0023 | **Direct topology constraint.** ODT remains the dedicated family owner; the change imports no concrete family or umbrella crate and moves no shared ODF capability. |
| 0024 | **Direct inventory constraint.** The current package topology and ODT semantic/XML ownership remain unchanged. |
| 0025 | **No direct change.** OGraph chart-area transactions are untouched. |
| 0026 | **No direct change.** Shared OLE directory metadata binding is untouched. |
| 0027 | **No direct change.** XLS sheet-anchor ownership is untouched. |
| 0028 | **Out of scope.** The ordered IWA monolith exit is unaffected; the task excludes iWork work. |
| 0029 | **Out of scope.** The archive-free IWA object-index foundation is unaffected; the task excludes iWork work. |

This was a documentation-only audit. No source file was edited, and no build,
benchmark, test, or commit was run by this review.
