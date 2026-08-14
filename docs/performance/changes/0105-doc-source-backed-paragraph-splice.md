# Change 0105: source-backed Word97+ paragraph splice

Date: 2026-08-14

Status: Accepted for narrow correctness and bounded selector coverage. This
record makes no end-to-end performance claim.

## Scope

Production commit `9bd292c4f` adds
`litchi_doc::body_text::source`, a separate immutable source-backed owner for
one ordinary main-story paragraph in a Word97+ binary DOC. The owner opens a
caller-provided positional `ReadAt`, validates the CFB/DOC envelope and FIB,
resolves a zero-based paragraph selector, and admits a replacement only when
the complete selected paragraph (including its terminating paragraph mark) is
one uncompressed Unicode CLX piece and the replacement has exactly the same
UTF-16 width. The public owner is intentionally separate from the existing
owned body editor; it does not rebuild the FIB, CLX, or FKP tables.

Paragraph selection scans Unicode pieces through bounded chunks (at most 4,096
UTF-16 units per request, subject to the configured finite prefix limit) and
retains only the selected paragraph bytes/text. A successful edit submits one
same-length `WordDocument` range splice to the validated CFB publisher. The
candidate is reopened and the selected paragraph, physical range, replacement
bytes, and target fingerprint are checked before a `Commit` is returned.
`Commit` exposes the content-free diagnostics, complete sequential publisher,
atomic path publisher, and a source-checked reversible `Patch`/`inverse`.

## Contract and refusal boundary

The snapshot captures source version, length, and a complete artifact
fingerprint. Open, selector resolution, planning, candidate validation, and
publication re-check source identity; stable-token mutation, stale/foreign
source, target-fingerprint, and candidate-readback mismatches fail before a
successful commit. An exact no-op returns the original immutable snapshot and
patch identity and streams the exact source bytes. Sequential sinks preserve
the existing typed partial-output behavior.

The closure refuses compressed or cross-piece paragraphs, empty paragraphs,
UTF-16 width changes, structural/control units, fields, drawings/objects,
selected revision properties, fast-save CLX, piece PRMs, encrypted or
obfuscated FIBs, protected documents, macro or signature owners, ambiguous
topology, malformed bounded ranges, and every configured finite-limit
overflow. Unsupported FIB/table layouts remain typed refusals rather than
falling back to a topology-changing writer.

## Verification

The source owner has 15 focused tests. They cover surrogate-preserving UTF-16
encoding, bounded geometric paragraph buffering, generated and repository DOC
fixtures, same-width replacement and inverse restoration, exact no-op identity,
typed width and partial-sink failures, later-piece addressing, finite-limit
refusals, malformed FIB/CLX and protection checks, stable-token mutation,
candidate readback, and a giant-piece scan whose source reads remain below the
full piece size. The default-feature and all-feature `litchi-doc` library
suites, warning/deprecation-denied all-target checks, strict all-target Clippy,
rustdoc, formatting, and independent adversarial review of the production
implementation also passed.

The complete artifact fingerprint and CFB validation/publication scans remain
mandatory at open, planning, candidate validation, and output boundaries. The
bounded paragraph selector therefore does not establish proportional
end-to-end latency, physical I/O, range-read, allocation, peak-heap/RSS,
cold-filesystem, high-latency, or real-producer performance. It is not broad
DOC CRUD coverage: tables, fields, drawings, revisions, cross-document copy,
structural edits, and other stories remain outside this owner.
