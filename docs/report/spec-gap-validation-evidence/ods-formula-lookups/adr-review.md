# Accepted ADR applicability review

Status: pre-implementation design review. The independent ADR reviewer read the
accepted records and amendments and found no conflict with the proposed ODS
lookup/reference batch. This does not certify production behavior or validation.

| Records | Applicable constraint | Planned evidence |
| --- | --- | --- |
| 0001, 0004 | Panic-free typed outcomes; concise contextual private owners | Invalid-argument tests, strict lints, no new public API |
| 0002, 0009, 0010, 0023, 0024 | Spreadsheet semantics stay in litchi-ods; container/detection remain below it | Dependency-boundary gate and manifest diff |
| 0003, 0006, 0007 | Read-only evaluation preserves package/formula/cache state; explicit providers and checked coordinates | No package-write path; typed provider/source-failure tests |
| 0005 | Finite work/storage/read budgets, borrowed data, cancellation, source fences, semantically invisible caches | Resource suite, projected-lazy regressions, read accounting and measured allocation |
| 0008 | Buildable phases and evidence before support claims | Frozen source, independent reviews, seven isolated gates, oracle/native and performance receipts |

ADRs 0012–0022 and 0025–0027 govern other formats; their format-specific
representations and native histories are not ODS evidence. ADRs 0028/0029
cover the IWA migration and object-index owners. Later Numbers formula-cache
amendments do not specify ODS evaluator semantics. ADR 0015 supersedes 0014
for OOXML core properties; ADR 0024 resolves historical topology wording.
The canonical ODF split remains ADR 0023, not the IWA record renumbered 0029.

The batch adds no ambient network/file access, runtime, parallel pool, unsafe
code, package mutation, dependency recalculation or result-cache publication.
Source IRIs remain inert; unsupported resolution stays a typed capability error.
New lookup case-folding has a private owner and must not change existing scalar
comparison or conditional-criterion profiles.
