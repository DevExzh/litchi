# Change 0381: DOCX source-backed glossary entry batch

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope

Change 0381 extends the source-backed glossary text owner from one selected
entry to a general nonempty caller-bounded batch within one glossary. The batch
resolves semantic selectors in canonical source order and accepts the existing
glossary identity forms, while alias duplicates, overlapping selections, and
duplicate paragraph intents fail closed. Only paragraphs selected by the batch
are eligible for replacement.

The batch performs one topology resolution/materialization and one inventory
pass for the common glossary owner. Aggregate selector, entry, replacement,
and output limits apply across the complete batch rather than independently
per requested item. Every replacement size is measured before replacement
materialization, and staging keeps at most one temporary wrapper at a time.

This tranche does not add glossary entry creation, deletion, rename, reorder,
or metadata/background edits. Cross-part or general-story batching, managed
editing, durable patch wire, and broad DOCX behavior remain outside the
surface.

## Architecture and behavior

`litchi-docx` remains the semantic owner. The batch binds one validated
glossary Part, its source fingerprint, and the selected paragraph intents.
Canonical source-order resolution makes the semantic plan deterministic;
alias-equivalent selectors cannot address the same entry or overlapping
paragraph more than once. Selected-only paragraph ranges are staged with
checked aggregate accounting before text materialization. Cancellation during
planning or staging leaves the transaction unchanged.

Publication changes exactly one glossary Part and preserves the glossary root,
entry properties, sibling entries, opaque XML, unrelated package members, and
package topology. The established exact no-op, source-freshness, source-bound
inverse, signature, and failure-atomicity fences remain in force. No semantic
interpretation of fields, revisions, controls, or rendering is introduced.

## Resource-accounting boundary

The batch accounts for all replacement sizes before materializing replacement
text and does not retain multiple temporary wrapper buffers concurrently. The
selector, entry, replacement, and output ceilings are aggregate boundaries
for the complete request. These are implementation and refusal boundaries;
they are not a fixed-memory or system-level OOM result.

## Verification

The focused glossary-batch integration suite passed `18/18`. It covers
nonempty single and multi-entry same-glossary batches, canonical source-order
resolution, alias duplicate and overlap refusal, selected-only paragraphs,
aggregate selector/entry/replacement/output limits, unique paragraph intents,
cancellation-atomic staging, replacement-size premeasurement, temporary
wrapper accounting, exact one-Part publication, no-op and source-bound
inverse behavior, source freshness, and root/sibling/opaque XML preservation.
It also covers the stated refusal boundary for unsupported lifecycle and
cross-owner operations.

The existing source-backed story-text suite passed `11/11`; the
default-feature library passed `926/926`; the all-features library passed
`935/935` and all integration binaries passed. DOCX doctests passed `74` with
`31` ignored. Strict Clippy passed with `-D warnings`. The crate-boundary gate
passed for 64 workspace packages and 240 internal dependency declarations
with 14 existing debt entries.

Validation used one Cargo process and one test run at a time,
`CARGO_BUILD_JOBS=1`, disabled incremental/debug build state, one dedicated
target, serial test threads, a 6 GiB per-process virtual-memory cap, and a
`>=10 GiB` available-memory launch threshold. This is an OOM-mitigating,
resource-capped procedure, not proof of OOM prevention.

## Claim boundary

`performance_claim: none`; `claim_authorized: false`. The evidence is limited
to the same-glossary source-backed batch lifecycle, its canonical selector and
alias refusal rules, aggregate resource boundaries, cancellation atomicity,
exact one-Part publication, no-op/source-bound inverse, freshness, and
preservation invariants. No benchmark, latency, allocation-volume, RSS,
physical-I/O, cold-cache, throughput, scaling, fixed-memory, broad DOCX,
cross-part/general-story, managed-editing, durable-patch, or system-level
OOM-prevention claim follows.
