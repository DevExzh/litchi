# 0350 — Owner decision: resolve the API deferrals (2026-10-20)

- Status: owner decision, recorded by the coordinator.
- Date: 2026-10-20.
- Overrides: owner memo [0331](0331-owner-decision-memo.md) items 14–20,
  23, 24, 26 and 29. Confirms every other item.

## 1. Decision

The owner wants one consistent API experience across every format crate:
the leaf crates and the unified `File` facade. The API deferrals of record
0343 (D1–D7, D10, D11), memo item 26 (`litchi::sheet::text`) and memo
item 29 (`Box<dyn>` in the `litchi-core::sheet` traits) are **required
engineering work, not accepted long-term exceptions**. They are to be
resolved promptly, foundational changes first, with every affected consumer
migrated in the same change, preserving correctness, fidelity and
performance.

| Memo item | 0343 | Overridden default | Required outcome |
|---|---|---|---|
| 14 | D1 | Direct `&mut self` package mutators kept | Immutable `Package` APIs: no public `&mut self` mutator on any `Package` (DOCX, PPTX, XLSB `into_workbook` and mutating `Workbook` verbs, and any other crate); mutation goes through `Edit` families; outputs take `&self` |
| 15 | D2 | Scoped OOXML commit triad kept | One Edit/Commit/Patch/Diagnostics contract: every scoped family commits to a `Package` with `Diagnostics`; no scoped `Transaction` names, no part-level `snapshot()` publication |
| 16 | D3 | PPT scoped `Transaction` families and `rollback` kept | PPT scoped families follow the same contract; `rollback` removed, the writers re-plumbed |
| 17 | D4 | `Error::UnsafeEdit`, ODT `MutableDocument` kept | One error taxonomy across crates replaces `UnsafeEdit`; ODT authoring follows the canonical `Builder`/`Edit` shape |
| 18 | D5 | PPT `raw::Reader` and 173 `*_with_limits` entry points kept | Configurable limits through the canonical options shape; no limits-only ingress forks |
| 19 | D6 | Flat ODF fixed ceilings and crate `Limits` kept | Flat ODF and ODC owners take caller-configured limits through `OpenOptions` |
| 20 | D7 | `Metadata.keywords: Option<String>` kept | Structured metadata: keywords as a list, with the JSON/CLI/Python shape migrated together |
| 23 | D10 | iWork `CellPosition`/`CellRange` deferred | Renamed to `At`/`Area` |
| 24 | D11 | `litchi-iwa` `*DocumentBuilder` prefixes kept | Canonical naming and module ownership for `litchi-iwa` |
| 26 | — | `litchi::sheet::text` kept in the facade | Moved to a leaf crate (module ownership), names aligned with ADR 0034 |
| 29 | — | `Box<dyn>` in `litchi-core::sheet` traits accepted | Static dispatch on the sheet traits; the per-cell evaluation-path boxing removed |

## 2. Deferred

Only memo item 22 (0343 D9), the comprehensive content-redaction audit of
error and `Debug` strings, is deferred. Its existing redaction
requirements still apply to every change made under this decision.

## 3. Confirmed

The owner accepts every other adopted scope, capability-reporting,
behavioural and performance decision of the memo: items 1–13, 21
(withdrawn), 25, 27, 28 and 30.
