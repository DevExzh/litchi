# Log sections for change 0756

## For `HOTSPOTS.md`

## 0756 — the OLE2/OOXML wave of 2026-09-22/23 integrated; the next hotspots

[0756](0756-ole2-ooxml-wave-integration.md) integrates records 0742–0755 and
0757. A descriptive before/after sweep of the wave (base `009d515bef`, tip
`ddf788eb80`) gives per-format geometric means of 0.675 (DOCX), 0.813 (PPTX),
0.551 (XLSX), 0.864 (DOC), 0.813 (PPT) and 0.502 (XLS). What leads now:
- XLSX commit+save is 60% deflate and 12.6% audit;
- PPTX per-capture notes-graph validation;
- DOCX's remaining full source audit at commit;
- XLS per-string UTF-16 temporaries;
- DOC fresh-write CFB zero-fill;
- PPTX streaming per-member deflate;
- incompressible `opc_mutated_save` (not yet profiled).

Six owner decisions are pending, including proposed ADR 0032. The program goal
remains open. [Evidence](results/change-0756/README.md).

## For `REPORT.md`

## 0756 — fifteen records, each changed by an independent review

[0756](0756-ole2-ooxml-wave-integration.md) summarizes the wave:
- **PPTX:** owned media cross-copy 416.7 → 121.8 ms (0742, 0751); full text and
  1% edit on a 100 × 100 deck about 1.85 times faster (0743).
- **XLSX:** dense-sheet first cell 28.4 → 7.8 ms and one-cell commit+save
  152.7 → 60.9 ms (0744).
- **XLS:** edits 1.5–9 times faster (0746, 0748).
- **DOCX:** no-op and one-edit saves 2.5 and 1.9 times faster (0754); streaming
  creation 1.9 times faster (0752).
- **Legacy fresh writers:** 2–3 times faster (0753).

Correctness fixes: ADR 0006 audit gaps (0750), a panic that aborted the
process (0755), an audited-handle gap (0754), nondeterministic XLS output and
silent truncation (0757). One memo was withdrawn under ADR 0005 and proposed as
ADR 0032. The sweep is descriptive (9-sample processes, one core, shared host).
Every flag above 5% is disclosed, with instruction counts where they separate
work from code layout. [Evidence](results/change-0756/README.md).

## For `GOAL_AUDIT.md`

## 0756 — wave integration: measured gains, reviewed correctness, open owner decisions

[0756](0756-ole2-ooxml-wave-integration.md) closes a coordinated wave.
- **Evidence:** Opus implementers in isolated worktrees, independent adversarial
  reviews (every branch changed as a result), identical-command before/after
  builds, and a wave-wide descriptive sweep.
- **ADR work:** accepted ADRs were kept as hard constraints. One
  ADR-incompatible memo was withdrawn and proposed as ADR 0032. A coordinator
  reading of ADR 0005's memo amendment is listed for confirmation.
- **Also listed for the owner:** G5, deterministic byte changes for creation,
  save durability and budget leases.
- **Still unproved:** physical cold-cache, remote/range, concurrency scaling,
  and full CRUD-category completion.

The non-iWork goal remains active. [Evidence](results/change-0756/README.md).
