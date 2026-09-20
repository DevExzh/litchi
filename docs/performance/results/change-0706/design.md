# 0706 candidate and constraints

The 0705 source-bound publication profiles put source XML auditing at
54.56–58.13% of publication guest instructions. `inspect_attributes` is a
material nested owner; duplicate bookkeeping includes range-vector growth.
Those nested costs overlap and are not additive speedup predictions.

The candidate reuses a result from the existing lexical layout traversal.
Only a successful zero/one-attribute proof suppresses duplicate-key checking.
The ordinary attribute iterator still parses the attribute and charges its
budget; `xml:space` still follows the same decoding and inherited-state rules.
For multiple attributes, the original checked iterator remains authoritative.
The lexical traversal and its failures still precede attribute inspection.

This differs from rejected 0529's extra unchecked probe followed by replay.
It introduces no extra probe, pass, retained collection, dependency or public
API. Its saturated count cannot overflow. Both slice and guarded stream
callers must use the proof from their existing successful layout check.
Declaration checks discard the proof.

Independent source review found that current XLSX layout facts cannot replace
OPC's generic XML audit. Existing `SourceXmlPart` construction would introduce
different source/final validation, budgets and error timing rather than remove
work from this one-shot workflow. The insertion-only splice proof is also not
a general value-rewrite proof. Both original and replacement audits remain.

| Constraint | Candidate obligation |
| --- | --- |
| ADR 0001 / 0006 | Preserve every accepted/rejected input, typed error, offset and validation order. |
| ADR 0005 | No hidden retention or execution; measure latency, allocation and adverse controls separately. |
| ADR 0011 / 0024 | Keep XML grammar ownership in `xml-minifier`; no facade/archive shortcut. |
| Existing source publication policy | Retain source-compatible spelling, all finite limits, original/replacement audits and exact no-op behavior. |

Root freezes binaries before applying the candidate. Native and allocator
regions are separate. Public-API differential oracles and consumer gates are
required; full correctness cannot be inferred from two performance shapes.
The frozen primary admission gate is intentionally the same magnitude as the
historical probe gate. If it fails, restore production and retain the evidence.
