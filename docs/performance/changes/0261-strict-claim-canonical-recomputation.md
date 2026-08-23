# Change 0261: strict claim canonical recomputation

## Status

The verifier hardening landed in
`c96233f38541231cb1c2b8e864ff64f44481f16d` and this record documents the
integrity and bounded-memory gate. It makes no latency, throughput, allocation,
RSS, or speedup claim.

## Problem found

The previous strict path still trusted summary-derived acceptance metadata after
checking the summary and manifest hashes. In particular, a modified summary
could replace the accepted-statistic labels and then rebind its hashes. The
resource verifier likewise accepted declared paired deltas: a report with raw
control/candidate values of `100` and `1000` could present a `2%` relative delta
and pass the `5%` resource threshold.

## Independent recomputation

Strict verification now validates every raw report and recomputes the canonical
summary from the four compressed ABBA legs. Each report is decompressed and
`_project_report` sequentially validates its raw samples, recomputes bounded
elapsed statistics and identity projections, and discards the raw report/sample
payload before the next leg; it retains no elapsed sample values. Only bounded
rows, identities, source and sink metadata, and the recomputed statistics
needed for the final summary remain. The elapsed p50/mean/p95/p99 values,
accepted/adverse cells, drift decisions, and complete canonical summary are
derived from those bounded recomputations; the summary's accepted-statistic
labels and result cells cannot define the verdict. The final comparison covers
the complete canonical summary, not just its result cells.

The raw-report profile is detected from report metadata before summary work.
The historical `legacy-v1` reports and additive `current-v1` reports each
reproduce their existing direct summaries exactly; a mixed four-leg profile is
rejected. Canonical identities are hashed incrementally so the verifier does
not create a second report-sized canonical byte buffer.

Resource verification derives each requested metric's A1/B1/B2/A2 values from
the parsed `heaptrack` and `/usr/bin/time` leg sources. It independently
recomputes control and candidate count/mean/median/minimum/maximum aggregates,
the `A1 -> B1` and `B2 -> A2` ratios and relative deltas, and each observed or
ineligible status before comparing the declared report. The `time` and
`heaptrack` run/status/artifact/parser identities are fail-closed. Each resource
leg binds exact A1/B1/B2/A2 variant, revision, binary, harness tool, and profile
metadata; missing, mismatched, or unsupported identity cannot yield a resource
result.

Public projection helpers are not exposed. The public verifier path creates the
module-private `_ValidatedProjection` trust carrier only after raw validation;
plain mappings are rejected by the projected summarizer, and mutations of the
carrier fail its integrity check before summarization. Raw projection-marker
fields are ignored when a report is treated as raw input. Only the verifier's
validated bounded projections may carry internal summary identity,
elapsed-statistic, or operation-metric values into the final summary.

## Bounded input and memory behavior

The streaming verifier fails closed at these limits:

- 512 MiB decompressed bytes per compressed ABBA member;
- 2 GiB total decompressed bytes across the four members;
- 64 MiB for the JSON summary.

After each `_project_report` call, the decoded raw report and sample payload are
discarded before the next leg is decoded; only its bounded projection is
retained. These are verifier input ceilings, not a claim about the library's
runtime memory use. An external `/usr/bin/time -v` run of the strict checker
recorded a maximum RSS of `1,114,076 KiB` for the retained four-report evidence
package.

## Verification

- Strict claim validation: `OK: 4 performance claims validated (strict)`.
- Relevant Python verifier and summary tests: 141 passed.
- Adversarial coverage includes raw sample tampering, elapsed-statistic and
  accepted-set tampering, legacy/current profile exactness, mixed-profile
  rejection, sink/identity tampering, raw resource-source absence, declared
  resource-only values, values-by-leg mismatch, paired-delta mismatch,
  time/heaptrack run/status/artifact/parser identity, exact resource
  variant/revision/binary/harness-tool/profile binding, ignored raw projection
  markers, fabricated/mutated `_ValidatedProjection` rejection, and
  decompressed-byte/process cleanup limits.
- Independent review confirmed the fail-closed resource and private-projection
  integrity boundaries and that the no-speedup claim boundary is unchanged.

This change strengthens the trust boundary for existing performance evidence;
it does not add a new claim-registry entry or alter any measured performance
result.
