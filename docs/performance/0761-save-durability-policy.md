# 0761 — explicit save durability policy

This change implements the explicit per-call durability policy authorized by
[0758, decision 5](0758-owner-decisions-2026-09-24.md), including its ADR 0005
amendment. The original implementation commits are `678945d928`, `f9265f57a2`,
`a1010f8519` and `9aeed778b9`.

The original branch's evidence draft is incomplete and is not a published
performance result. Current integration, validation and limitations are recorded
in [0773](0773-save-durability-integration.md). No old scratch measurement or
missing gate artifact is adopted by this record. `performance_claim: none`.
