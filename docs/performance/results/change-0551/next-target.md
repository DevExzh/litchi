# Next: implement the compact source-cell proof and differential oracle

Use the measured source-backed `MultiSourceEdit` path. The compact proof must
remove repeated whole-sheet cell-address/tag reconstruction as well as a
second XML reader; merely attaching the full Scanner to planning leaves its
per-cell work and adds retained state. Avoid row-only reparsing because the
dense 1% workload edits every dense-sheet row. These are design selections,
not measured rejection of an implemented candidate.

Start with a private proof builder in the worksheet owner, supplied original
event start/end positions from the already eligible shared traversal. Reuse
the raw parser's resolved row/cell addresses only through an explicit private
handoff whose event timing also covers empty cells; do not infer the address
by inspecting a last cell that might belong to the preceding event. No raw
parser or source validator semantics may be bypassed.

A builder refusal must disable and drop only the builder. In particular,
do not return false from the existing observer just because layout proof is
unavailable: that boolean currently requests authoritative source reparse.
Preserve the original validator/parser success path and exact diagnostic order.
Only publish proof after validator EOF, successful raw finalization and all
scanner-equivalence checks. This includes decoding/normalizing the attributes
which the scanner would inspect on unchanged tags. No provisional error may
become a new public planning error.

The representation should retain source-bound cell and row spans plus the
minimal global facts needed for existing-row scalar value edits. Prove each
omitted `Layout` field through validated absence, an equivalent check, or
explicit authoritative fallback; do not supply guessed defaults. Typed and
prefixed cells in the noncompact shape must be included in the candidate
campaign. Root/sheetData/row namespace context and dimension expansion remain
obligations even when only a single value changes. Discontiguous primary spans
and unsupported/shared formula or membership-changing cases may decline the
proof route while retaining the current complete rewrite.

Before timing, compare complete output bytes and exact errors against the
unchanged scanner/writer for accepted and rejected source variants. Exercise
ordering, inferred addresses, bounds, empty/start/end forms, namespaces,
unused declarations with invalid values, attribute normalization, dimensions,
formulas, failed proof allocation, limits/cancellation, no-ops, multi-sheet
atomicity, cloned snapshots, inverse patches and source mismatch. The source
audit in `design-review.md` is the checklist; it is not this oracle.

Freeze metadata byte/event caps and an unchanged complete fallback before the
matched native/allocator campaign. Measure planning, staging/commit,
publication and retained state across the commit/patch lifetime, with managed
execution and refusal/cap controls. Keep full output validation and independent
changed-cell readback. Retain the candidate only on representative workflow
and memory gates; record and revert failed speculation. If this architecture
cannot satisfy those conditions, return to the measured XML-validation owner.

OLE2 and OOXML optimization remain the goal. ODF waits until that goal is
complete; iWork remains outside this workstream.
