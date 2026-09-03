# Change 0378: DOCX source-backed secondary-story text

Status: implemented

performance_claim: none

claim_authorized: false

## Scope

Change 0378 adds the missing bounded source-backed text lifecycle for selected
entries in DOCX footnotes, endnotes, and comments. The selectors can read the
selected entry text through a bounded snapshot and stream it to a caller, or
stage an exact text replacement for source-bound publication. Main-document,
header, and footer stories retain their existing scope; this change adds the
secondary-story entry owners without broadening the API to arbitrary XML.

Glossary entry text is deliberately not included. The existing glossary
catalog and auxiliary-part lifecycle remains available, while glossary entry
text is deferred to a subsequent reusable-engine tranche.

## Architecture and behavior

`litchi-docx` remains the semantic owner. Entry selection scans and splices the
exact raw body range for the requested footnote, endnote, or comment. It
validates unique relationship and content-type ownership, inbound ownership,
the Strict or Transitional package dialect, package-wide story edges, and the
current source version before exposing or changing bytes. Missing, ambiguous,
reserved, managed, unsupported, or unsafe topology returns a typed refusal.

The snapshot and transaction paths retain the existing exact no-op, inverse,
signature, stale-source, and failure-atomicity fences. Candidate replacement
validation completes before counters or staged state are assigned. Namespace
contexts structurally share prefix and value bytes, and text/reference
decoding is bounded and chunked rather than converting unbounded source text
before the output limit. Entry, entry-byte, namespace, and output limits
remain enforced at their respective boundaries.

Unrelated XML, root attributes and prologs, relationship parts, package
topology, and unsupported content remain preserved. Managed edits and semantic
interpretation of fields, revisions, controls, or rendering are not added by
this change.

## Verification

The focused secondary-story integration suite passed `25/25`, covering
Transitional and Strict footnote, endnote, and comment selection, text
publication, exact preservation, missing/ambiguous/reserved IDs, relationship
and dialect topology, managed and cancellation refusal, limits, legal and
unknown references, output accounting, signed no-op/change behavior,
stale/foreign/source mutation, and sink behavior. The existing story-text
integration suite passed `11/11`; the default-feature library passed
`926/926`; the all-features library passed `935/935` and all integration
binaries passed. DOCX doctests passed `74` with `31` ignored. The
crate-boundary policy passed.

Validation used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, disabled incremental/debug build state, one dedicated
target, serial test threads, a 6 GiB per-process virtual-memory cap, and a
`>=10 GiB` available-memory launch threshold. This is an OOM-mitigating,
resource-capped procedure, not proof of OOM prevention. Strict Clippy passed
with `-D warnings`.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. The evidence is limited
to the selected secondary-story text lifecycle, its exact source/preservation,
refusal, atomicity, signature, inverse, and bounded-resource invariants. No
benchmark, latency, allocation-volume, RSS, physical-I/O, cold-cache,
throughput, scaling, fixed-memory, broad DOCX, glossary, or system-level
OOM-prevention claim follows.
