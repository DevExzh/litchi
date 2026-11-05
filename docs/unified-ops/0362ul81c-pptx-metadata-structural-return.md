# 0362ul81c — PPTX metadata structural finite review return

- Date: 2026-10-25.
- Status: source-only finite review return; typed-limit integration and Cargo
  validation remain with the coordinator.
- Base: `5170479a57e1065625fd210473ab1a6502012da2`.

The structure parser now rejects DTD and processing-instruction events while
they are still borrowed, before `into_owned` can allocate a rejected payload.
It also disables quick-xml text trimming in this graph-only parser so raw
whitespace event lengths are admitted before the selected string limit; text
remains semantically ignored as before.  Focused tests cover oversized
whitespace, DTD and processing-instruction payloads.

Guide replacement planning now uses the target index as the tie-breaker when
slide and notes entries share an insertion offset.  A regression covers both
targets absent from one existing common `extLst` and verifies slide-before-
notes ordering.

No Cargo, rustdoc, clippy, native build or ABBA run was performed.  The common
61 physical scan/no-op base check and record-84 typed-limit migration remain
separate coordinated work.
