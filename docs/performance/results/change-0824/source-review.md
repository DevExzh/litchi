# 0824 source and measurement basis

The base is `7b268927bfe3ed9c9000bfd567fc4ce656ec0cd9`. Its production and
ordinary-save harness sources are unchanged from the 0822 profile base; 0823
rejected its transport rewrite and restored production exactly. The preceding
0823 seal replays against HEAD. All 35 normative input hashes and the three
unrelated workspace files match their preceding records.

## Mechanism to test

`Transaction::set_shape_text` reads the current complete Scene, rewrites the
selected raw span, reads the staged complete Scene, checks shape count and
selected text, installs the payload, and records its allocation through a
`Weak`. A replaced payload misses that proof. `compact_changed_slides` already
uses it to avoid reparsing unchanged staged bytes, but if compaction changes
even one byte it reads both complete Scenes and compares shape IDs, names,
and text. A proof miss still reads the staged Scene before compaction.

In the pinned real fixture, after the exact selected text replacement the
published slide differs only by deletion of CRLF between the XML declaration
and document root. `fixture_basis.py` derives and checks that byte difference
from independently hashed source/reference ZIP members. It is one-file byte
evidence, not a universal semantic or timing result.

The candidate records source/output document-root ranges during the existing
compactor event pass and compares their complete bytes. Invalid/unavailable
ranges conservatively miss. A match permits skipping the repeated comparison:
the staged Scene has already succeeded, the complete root bytes are identical,
and the unchanged compactor outside that root preserves BOM/declaration/PI/
comment bytes while only removing ASCII whitespace. Other outside content is
still refused by the existing paths. It changes no emitted byte. A root change
retains the old comparison. Debug builds rederive the semantic equivalence.

The intended counted-read effect for compaction that only changes outer
whitespace is two to zero on a valid staged-allocation proof, or two to one on
a miss. Byte-identical compaction retains its existing behavior. No new cache,
memo, retained payload, or transaction field is proposed; ranges and a boolean
are operation-local constant-size state.

## Existing profile evidence

The exact-owner 0822 frame census has 2,971/2,996 samples. Scene occurs in
636/624; `compaction_scene` is present in 301/304 Scene-qualified stacks.
The independent `profile_basis.py` checks the retained compressed and decoded
frame hashes and exact executable identity. Counts are overlapping inclusive
stack presence, not additive time fractions, exclusive cycle costs, or an
Amdahl estimate. The new trial must measure public workflow latency and
allocations before adopting anything.

The initial fingerprint still requires exact logical payload digests under the
existing revision/source-authorization contract. Replacing those with CRCs or
compressed bytes is not equivalent. Existing digests and slide-root verdicts
already reuse unchanged allocation identities. A variable-size Scene cache
would require ADR 0005 eviction/budget handling; this candidate does not add one.

## Invariants and gates

- ADR 0003: original staged validation, dependency closure, commit capture,
  source authorization, reversible patches, and atomic publication remain.
- ADR 0005/0032: no new cross-operation state or parsed-value retention; no
  hidden executor or ambient runtime behavior; fresh matched measurements.
- ADR 0006: identical output bytes, unknown/MCE content and namespace spelling,
  initial refusal ordering, protection/signature policy, and finite limits.
- ADR 0001/0004: no public API or dependency change.
- ADR 0010/0011/0024: transaction/XML ownership stays in the PPTX owner.

Required tests prove the saved calls and compare baseline/candidate behavior
for BOM, outer whitespace, empty roots, MCE/unknown content, inner whitespace,
attribute normalization, invalid outer content, and malformed/limited inputs.
The default parser ceilings remain unchanged; shorter outer whitespace cannot
increase any structural count or retained semantic text. Byte offsets are
recreated by the existing final capture, not reused across changed buffers.

Both legs require complete PPTX quality gates and all nineteen workflow output
oracles. Native and allocation lanes run separately, in frozen counterbalanced
orders. The adoption threshold is unchanged in form from 0823: at least one
eligible paired p50 gain of 3% with upper interval below one, no lower interval
above 1.05 on any row, and no allocation-resource increase. No outcome is
claimed before the admitted captures and independent readers complete.
