# 0784 next candidate: known notes namespace fast path

This is a bounded, read-only source review. It records a possible follow-up
for `notes::resolved`; it does not change production code or claim a speedup.

## Candidate and profile qualification

The seam is [`crates/litchi-pptx/src/notes/mod.rs:78`](../../../../crates/litchi-pptx/src/notes/mod.rs:78),
where `resolved` currently converts every bound namespace with
`std::str::from_utf8`. The 0784 frame-pointer follow-up has the exact capture
owner on 1,162 of 3,287 and 1,153 of 3,286 total sample stacks in its two
retained large runs. `notes::resolved` appears beneath that owner in 146 and
150 samples. These observed counts support investigating the function; they
do not establish a production latency share or speedup. The coordinator's
final qualification review refuses phase-fraction claims under the frozen
plan because one recovered capture stack has an unresolved interior frame.

## Safe shape

The existing notes module already owns the six exact PresentationML namespace
constants:

| value | constant | byte length |
| --- | --- | ---: |
| Transitional PresentationML | `P` | 58 |
| Strict PresentationML | `PS` | 46 |
| Transitional DrawingML | `A` | 53 |
| Strict DrawingML | `AS` | 41 |
| Transitional relationships | `R` | 67 |
| Strict relationships | `RS` | 55 |

For `ResolveResult::Bound(Namespace(value))`, compare the byte slice against
these exact constants and return the corresponding static `&str` on a match.
For every non-match, retain the current `std::str::from_utf8(value).map_err(xml_error)` fallback. A length-dispatched or length-guarded comparison is preferable to six unconditional equality checks: the six lengths are distinct, so a known URI pays for one exact comparison and an unrelated URI usually avoids all full comparisons. The `Unbound` and `Unknown` arms should remain byte-for-byte and error-for-error unchanged.

The returned static strings are valid for the current `&'a str` result lifetime:
`'static` outlives `'a`. Callers only compare the returned namespace and use its
length; they do not rely on allocation identity. Exact matching is
case-sensitive and does not alter URI semantics. No parser event, MCE branch,
XML budget, attribute count, attribute-byte count, or validation order moves.

The fallback is essential. A valid but unrelated namespace must still be
accepted as the same borrowed string, and an invalid UTF-8 namespace must still
produce the same `Error::Xml` from `from_utf8`. This fast path must not treat
near misses, different lengths, empty bound namespaces, or vendor namespaces
as one of the six known values. `ResolveResult::Unknown` must continue to
produce the existing `Invalid("unbound XML prefix ...")` result.

## Code-generation and measurement risk

The optimization replaces UTF-8 validation of the six known ASCII URIs with an
exact byte comparison and a static return. It may help because the profile
shows `from_utf8`/namespace conversion repeatedly inside the notes scanner,
but equality is still a byte scan. A naive chain can add branches and repeated
comparisons for `A`, `AS`, `R`, `RS`, and unknown values, or increase inlining
code size. The implementation should be checked in release assembly or a
paired native profile before claiming benefit. Keep the function small and
avoid a lookup table or allocation in this hot path.

## Required tests before implementation is accepted

1. Unit coverage for all six exact constants, plus an unrelated valid UTF-8
   URI of each known length and near-miss values, asserting the same returned
   text as the current implementation.
2. Direct `ResolveResult` coverage for invalid UTF-8 bound bytes, an empty
   bound namespace, `Unbound`, and `Unknown`, asserting the exact existing
   success/error values and messages.
3. The existing notes scanner differential corpus over Transitional and Strict
   roots, including vendor namespaces, MCE-processed bytes, malformed prefixes,
   invalid names/values, duplicate attributes, and all XML limits. Compare
   both `XmlScan` projections and refusal results against the pre-change
   oracle.
4. End-to-end opened-capture checks for notes graph classification, revision,
   published bytes, and semantic identities. The ordinary no-notes slide
   validation path must remain unchanged.
5. A matched native capture measurement using the existing 0784 owner-qualified
   wrapper, with source/output/error oracles retained. Report the qualified
   `resolved` share and capture p50 only after the paired control/profile
   result passes its perturbation and identity checks.

No new memo, retained payload, or accepted ADR is needed for this candidate;
it is safe only as a local representation optimization if the fallback and
test matrix above remain exact.
