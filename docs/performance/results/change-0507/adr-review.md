# 0507 ADR and semantic review

All 30 ADR input files, including the README and 29 numbered records, are
byte-identical to the accepted-ADR inventory reviewed for 0506. Root verified
every file against that manifest and retained a fresh revision-bound
[manifest](adr-manifest.json). The [0506 applicability matrix](../change-0506/adr-review.md)
remains applicable; this record adds the value-decoding-specific obligations.

| Contract | Application to this change |
| --- | --- |
| 0001, 0004: correctness, typed API, semantic layers | Helpers remain private; public selectors, result types and ownership do not change. |
| 0002, 0009–0011, 0023–0024: dependency/grammar ownership | ODG owns these XML fields; no manifest edge, container API or facade export changes. |
| 0003: immutable snapshots and atomic publication | Batched values are transient parser locals; no partially decoded state is published. Existing edit, patch, no-op and candidate-readback paths remain in force. |
| 0005: bounds and measured performance | Fixed compile-time request arrays replace repeated work. Production groups each request eight unique fields. No cache, thread, unsafe code, ambient provider, configured-limit change or source-ownership change is introduced. Fresh before/after and whole-child profiles gate acceptance. |
| 0006: preservation and validation | Same namespace matching, normalized value decoder, expanded-name duplicate check and raw checked iterator are retained. Raw source XML and lexical span selection remain authoritative for edits. |
| 0007–0008: models and verification | No model feature or native compatibility claim changes. Owner tests, malformed cases, exact edit/inverse coverage, lint/docs/downstream and boundaries are scoped verification evidence. |
| Other accepted format/domain records | No touched owner, semantic vocabulary or publication rule; their detailed dispositions remain as in the unchanged 0506 matrix. iWork production work is outside this batch. |

Two validation boundaries constrain the implementation. The geometry and line
geometry requests may be combined only after the existing z-index lookup and
before the existing lexical-validation call. The second group may combine
only the eight string property lookups evaluated after lexical validation.
Conditional frame/3D reads, the name lookup, z-index conversion, and viewBox /
points lookups used as validation arguments retain their original positions.

A batch visits attributes in source order, which can discover a different
error first than request-ordered scalar lookup. Any raw iterator, decode, or
duplicate-expanded-name error must discard partial batched values and replay
all requests through the original helper in request order. This retains both
the exact decoder-before-duplicate behavior and errors in later raw attributes.
No error result is persisted or used to bypass retry under a distinct source.

Success moves the decoded strings into the original parser fields. The fixed
arrays introduce no user-controlled capacity or persistent state. Large values
remain under existing source admission; exceptional input can perform one
additional bounded attribute pass before the original failing path. Full
resource-allocation failure recovery is not newly claimed.

Independent source review and private/public differential tests verify these
obligations. The final gate receipt, performance summary and review decision
record observed results; this design matrix alone does not prove performance
acceptance, full-goal completion, or native-application interoperability.
