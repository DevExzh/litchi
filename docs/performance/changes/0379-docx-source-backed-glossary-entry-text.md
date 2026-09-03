# Change 0379: DOCX source-backed glossary entry text

Status: implemented

performance_claim: none

claim_authorized: false

## Scope

Change 0379 adds the missing bounded source-backed text lifecycle for one
existing DOCX glossary entry. Callers may select an entry by unique
Unicode-caseless name, canonical glossary ID, combined name and ID, or checked
source-order index. A snapshot can project or stream selected text, and its
edit can stage exact paragraph-text replacements for source-bound publication.

This does not add glossary entry creation, deletion, rename, reorder,
properties, categories, behaviors, background, or multi-entry text editing.
The existing glossary catalog and auxiliary-part transaction remains the
owner of its documented metadata and part-lifecycle operations. Managed entry
editing, MCE branch selection, fields, revisions, controls, layout, and
rendering remain outside this tranche.

## Architecture and behavior

`litchi-docx` remains the semantic owner. A borrowing `quick_xml` scan checks
the glossary root, `docParts` collection, each `docPart`, optional properties
before exactly one `docPartBody`, and the direct paragraph ranges of the
selected body. Name matching follows the glossary catalog's Unicode-caseless
semantics. Semantic selectors require exactly one match; index selection is a
checked snapshot-local source-order fallback.

Selection validates the unique main-document-owned internal glossary
relationship, exact glossary content type, inbound closure, auxiliary story
relationships, Strict or Transitional package/XML agreement, namespace and
relationship attributes, VML restrictions, XML name values, source version,
and finite selector, XML, event, depth, entry, namespace, paragraph, output,
replacement, and replacement-text limits. Input-dependent buffers use checked
size arithmetic and fallible reservation. Missing, ambiguous, orphaned,
external, wrong-owner, mixed-dialect, malformed, stale, foreign, managed,
cancelled, signed-change, or unsupported structures fail closed with typed
errors.

The resulting patch binds the resolved part name, exact body span, and source
fingerprint; publication does not re-resolve a name or ID that could have
changed. No-op and inverse behavior reuse the established story-text
transaction fences. Changed publication preserves glossary entry properties,
sibling entries, root metadata, comments and opaque XML, relationship parts,
unrelated package members, and the source artifact. Stale and signed-change
refusals complete before writing output bytes, while signed no-ops remain
byte-exact.

## Verification

The focused glossary integration suite passed `12/12`. It covers all four
selectors in Transitional and Strict packages, exact replacement and reopen,
no-op and inverse identity, source preservation, missing and ambiguous
identities, topology and dialect diagnostics, legal and unsupported entity
references, selector/entry/namespace/output/replacement limits, exact and
one-under namespace-wrapper boundaries including `&quot;`, managed and
cancellation refusal, signed no-op/change behavior, stale and foreign source
handling, zero-output refusals, and partial-sink errors.

The existing source-backed story-text suite passed `11/11`; the
default-feature library passed `926/926`; the final all-features library passed
`935/935` and all integration binaries passed. DOCX doctests passed `74` with
`31` ignored. Strict all-features library/test Clippy passed with
`-D warnings`. The crate-boundary policy passed for 64 workspace packages and
240 internal dependency declarations with 14 existing debt entries.
Independent API/compile and safety/test reviewers accepted the bounded change.

Validation used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, disabled incremental/debug build state, one dedicated
target, serial test threads, a 6 GiB per-process virtual-memory cap, and a
`>=10 GiB` available-memory launch threshold. This is an OOM-mitigating,
resource-capped procedure, not proof of OOM prevention.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. The evidence is limited
to this selected glossary-entry text lifecycle and its source identity,
preservation, topology, dialect, refusal, atomicity, signature, inverse, and
bounded-resource invariants. No benchmark, latency, allocation-volume, RSS,
physical-I/O, cold-cache, throughput, scaling, fixed-memory, broad DOCX,
multi-entry, or system-level OOM-prevention claim follows.
