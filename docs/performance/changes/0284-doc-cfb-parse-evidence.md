# Change 0284: DOC CFB parse evidence

Date: 2026-08-26

Status: Accepted deterministic instrumentation

Performance claim: none

## Decision

The DOC body-text profiling seam now exposes an operation-local, content-free
event pair for the one top-level in-memory `OleFile::open` used by strict DOC
ownership.
`Started` is emitted immediately before that parse and `Finished` carries
`Success` or `Error` immediately after it returns. The CFB observer is separate
from the existing high-level `DiagnosticEvent` observer, so the established
flat semantic phase stream remains unchanged. The callback is supplied by the
caller, is synchronous, and is never retained. The instrumentation uses no
global state, clock, or ambient runtime state; the observer receives no source
content, stream names, offsets, or physical identifiers. Observer panics
propagate, so pair balancing assumes every callback returns normally.

The public `Snapshot::open_bounded` and ordinary commit paths retain their
existing behavior and do not emit diagnostics. The feature-gated profiled path
performs a bounded source-length preflight before strict CFB parsing but emits
no high-level phase for that preflight. When package/source preflight and
strict-owner input validation permit parsing, the observer emits exactly one
pair around the top-level `OleFile::open`. A preflight failure therefore
produces no CFB parse events; a malformed CFB closes its parse event with
`Error`; strict semantic or public-reader failures after a successful
in-memory parse preserve their existing high-level phase ordering.

Changed profiled commits reopen the rendered candidate once for the strict
owner/public-reader validation chain. An exact byte-for-byte no-op commit
reuses the source snapshot and emits zero new CFB parse events during that
commit call; earlier edit staging is outside the observer window.

## Scope and claim boundary

These are deterministic phase and event-order claims only. This change
authorizes no latency, throughput, allocation, RSS, physical-I/O,
decompression, copy, or end-to-end performance claim. It does not change DOC
semantic behavior, source limits, preservation policy, or ordinary API
signatures. The observer is an opt-in diagnostics seam and is feature-gated in
both `litchi-ole-common` and `litchi-doc`.
