# ADR compliance: rejected fixed-markup candidate

| Constraint | Candidate and final state |
|---|---|
| 0001 priorities | Correctness precedes speed; extra complexity reverted for insufficient measured benefit. |
| 0002 / 0024 ownership | Fixed ODP grammar stayed in litchi-odp; no new dependency or crate edge. |
| 0003 immutable publication | Existing-document edits, commits and patches were untouched. |
| 0005 budgets and evidence | Explicit caller contexts; bounded stack scratch; exact limit fallback and ancestor rollback; source/binary-bound ABBA evidence. |
| 0006 preservation and validation | Exact fresh output and sink identity; common XML and publication validation retained; sink errors never retried. |
| 0008 verification | Candidate and restored baseline test evidence retained; no new native or platform coverage claimed. |
| 0010 / 0011 physical ownership | Common ODF/ZIP output ownership and accepted-byte progress unchanged. |

No accepted ADR change was needed. The candidate's bounded cancellation
completion rule was documented and reviewed. The final crate sources are
identical to the baseline, and the rejected candidate is evidence only.
