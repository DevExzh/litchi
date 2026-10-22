# 0360a — Return XLSX resource typing for correction

- Date: 2026-10-21.
- Reviewed candidate: `uo2/w4-xlsx` at `9c91b18f1475395a135bf2f90f489a8ac2b987aa`.
- Relevant earlier change: `01b9a9e25e` (typed limits and source replay).
- Contract: [0351 §2.4](0351-api-contract.md#24-error-taxonomy-item-17).
- Status: returned to W4; required correction and validation remain open.

The public presence of a `LimitKind` enum does not establish correct resource
typing. Review of the actual producers found string-based classification and
invented observations in the candidate. This is required engineering work,
not an accepted exception or part of the deferred redaction audit.

## Findings

`external_links::limit`, `ole_objects::limit`,
`rich_values::refresh_intervals::limit`, and the Data Model helper select a
kind by searching a resource description for words. They then report
`maximum + 1` even where the caller already knows the actual length or count.
For example, external-link serialization knows the complete `sheet_names`
length, but discards it when constructing its limit error.

There is a concrete dimension error in `ole_objects/codec.rs`: the combined
condition `nodes > MAX_NODES || stack.len() >= MAX_DEPTH` returns
`limit("XML structure")`. That description selects `LimitKind::Nodes` even
when only the depth bound was exceeded. The error consequently reports the
node ceiling and its fabricated successor instead of the observed depth and
depth ceiling.

`connections::invalid` treats any message containing `limit` or `exceed` as
a resource error, then guesses its kind and ceiling. The namespace binding
check knows `bindings` and `self.limits.max_bindings`, but passes only
`"connections namespace binding limit exceeded"`. The wrapper classifies it
as namespace **bytes** and substitutes `MAX_XML_BYTES`. Namespace depth and
metadata checks likewise lose their actual configured values. This approach
can also turn malformed-input overflow messages into unrelated resource
failures. `connections::embedded_data::limit` discards its typed `Resource`
argument before invoking the string classifier. Form-control resource
wrappers need the same audit.

## Required repair and evidence

Each resource-producing check must pass an explicit typed dimension and its
actual observed and maximum values. Split combined checks when their
dimensions differ. Preserve malformed-input classification, upstream typed
resource information, allocation errors and the content-free error contract.
A lower-bound observation after arithmetic overflow must be explicit and
justified; it must not replace an exact value already available at the check.

Add focused checks for depth versus node exhaustion, namespace binding count
versus bytes, caller-selected ceilings, oversized collections by more than
one item, and malformed inputs whose messages previously triggered the string
classifier. Check host error conversion retains the resource kind and maps
the error to `limit_exceeded`. Finish the raw/record and option-policy audit;
the inventory's zero rows cannot prove these semantic properties.

W4 owns the repair and its consumers. Existing performance captures remain
historical evidence. Final native acceptance must identify the repaired
candidate and retain all rejected attempts. The final integrated gate and
whole-workspace inventory remain required after the pending waves merge.

## Exact-replay review clarification

A separate review initially proposed adding a signed/encrypted policy guard
unconditionally to cross-owner exact restoration. That proposal is withdrawn:
changed authoring already refuses those sources, while exact replay checks
the complete artifact and therefore preserves its signature/encryption status.
Signed or encrypted no-op patches must remain replayable onto an independent
open of the same bytes. W4 is adding focused coverage for exact no-op copies
and for changed unsigned patches whose selected parts match a signed recipient
but whose complete artifact differs; the latter uses the ordinary guarded
replay path. No new refusal policy is authorized by this review.
