# 0837 — retryable ZIP writes and rejected fresh OPC compression

The retained production change fixes retryable sink interruptions in the
bounded owned ZIP entry adapter. The attempted fresh-only OPC compression
optimization is rejected at correctness qualification. No candidate release
build or comparative performance capture was admitted; no speedup is claimed.
iWork is excluded. The [packet](results/change-0837/README.md) retains the
candidate source, failures, baseline probe and exact diagnostic archives.

## Investigation and decision

[0836](0836-opc-filesystem-save-profile.md) localized observed source-save CPU
samples to Deflate, but did not authorize changing opened-package compression.
No bounded byte-transparent improvement was identified: original-target decode
still supports exact no-op and malformed-source checks; preservation already
feeds complete payloads to its byte-compatible encoder. Earlier 0618 evidence
also makes cross-member pooling unattractive.

Owner decision [0758 #4](0758-owner-decisions-2026-09-24.md) permits cheaper
compression for fresh creation while retaining opened-package rules. The
candidate routed newly authored `OpcPackage` publication through the existing
owned-entry level-5, 16 KiB staged protocol from 0762. The retained
`source_ingress` marker distinguished fresh packages from borrowed, owned and
source-backed materialized opens. Exact-source passthrough, validation,
publication planning and opened-package compression were unchanged.

That boundary is insufficient. The unchanged DOCX
`publication_raw_copies_unselected_members_and_inverse_restores_exact_artifact`
test fails at `source_backed_paragraph_copy.rs:183`: the durable inverse
restores the selected XML through opened-package compression, which cannot
reproduce the new fresh-source compressed stream. The fixture was created
with the candidate and then opened for source-backed editing.

Independent Python ZIP readback of both byte vectors retained by the assertion
confirms identical member names and decompressed payloads. Only the compressed
`word/document.xml` differs: 135 bytes in the fresh source, 141 bytes after the
durable inverse. Content types, package relationships and the unused binary
member retain identical compressed spans. The two reconstructed ZIP artifacts
and their hashes are preserved in `diagnostics/` and
`inverse-failure-analysis.json`.

The in-memory inverse can retain and copy the original artifact. A serialized
inverse replayed after reopen has a different obligation; XML restoration alone
does not prove exact physical restoration. The test remains unchanged. The
fresh-only dispatch, accessor and candidate OPC tests are reverted, with their
complete source retained for review. No compression level, staging boundary,
opened-package rule or durable patch encoding changes in production.

## Retained correctness fix

The candidate's short/Interrupted sink test first exposed a separate defect.
The lower owned ZIP compressor retains partially drained output and accepts no
new input when `write` returns `Interrupted`. The outer
`StreamingArchiveEntry::write` nevertheless poisoned the entry, so the next
`Write::write_all` retry failed permanently.

The adapter now refreshes output progress and propagates `Interrupted` without
poisoning. Hard sink failures, `WriteZero`, limits and malformed progress retain
their existing handling. No codec call boundary changes on successful writes.
A direct regression drives both Store and Deflate through short alternating
interrupted writes, checks uncompressed accounting at each retry, finishes the
archive and validates the full payload. Explicit `flush` retry behavior is
outside this write-only fix.

## Evidence scope and verification

The frozen experiment planned six cases, six paired blocks and 1,440 measured
operations. Only six baseline preflight reports / twelve measured operations
were collected, with one warmup per report. They cover tiny fresh XML, 2 MiB
XML, 256 small XML parts, 4 MiB pseudorandom data, and borrowed/owned edited
controls. They are diagnostic baseline runs, not a before/after result.
Package construction, validation, output hashing and destruction were outside
the prepared-graph `to_bytes` timer. Whole-process RSS includes preparation and
verification. The plan, probe, aligned lockfile and baseline executable identity
remain reproducible, but the comparative matrix was never admitted.

The rebuilt default-feature suites pass **6,847 tests, zero failures and 38
ignored tests** across 284 result groups. Final verification is recorded in `final-v2-quality.json`,
`red-green.json` and `closure.json`. The final gate set covers workspace
formatting, all-target compilation and warning-denied Clippy for ZIP, OPC,
DOCX, XLSX and PPTX; warning-denied rustdoc; their full default-feature tests;
and crate boundaries. This is not a full workspace/all-features or native
Microsoft Office interoperability claim.

Retained unsuccessful steps include the initial probe compile, malformed patch
context, a mislabeled baseline-only focused test, the candidate interruption
failure, and the durable inverse failure. Restoring OPC with `copy2` initially
preserved old mtimes and let Cargo reuse candidate artifacts: the entire first
`final-*` sequence is invalid as final verification. The corrected `final-v2-*`
sequence refreshes restored-source mtimes and rebuilds affected crates. No
failed measurement is replaced by a favorable rerun.

Offline closure verifies 37 recorded command receipts and the final two-file
production delta. Both exclusively owned temporary roots were removed: 20,289
files totaling 6,275,196,049 logical bytes. Closure and the independent ZIP
diagnostic replay pass again after cleanup. Unrelated workspace artifacts are
unchanged.

## ADR compliance and next work

| Constraint | Evidence |
|---|---|
| ADR 0001 / 0006: preservation and correctness first | Rejects the compression candidate; exact durable-inverse assertion unchanged. |
| ADR 0002 / 0010 / 0011 / 0024: ownership | Retained implementation is wholly in the ZIP adapter; no facade/archive dependency added. |
| ADR 0003: source-checked reversible patches | No patch representation or source check changes; physical inverse incompatibility blocks optimization. |
| ADR 0005 / 0031: explicit execution and bounded work | No threads, provider, I/O policy, buffer size or execution budget changes. |
| ADR 0008: verification | Final affected-owner gates and the retry regression remain inspectable with failure receipts. |
| Owner decision 0758 #4 | Fresh byte changes are conditional on preservation and durable-patch proof; this candidate fails that proof. |

The next compression attempt needs an explicit physical-restoration
representation or demonstrated compression provenance, tested through fresh
creation → source-backed edit → serialized inverse → exact archive recovery.
Changing opened-package compression to match new fresh bytes would violate the
current decision boundary. The broader non-iWork performance goal remains open.
