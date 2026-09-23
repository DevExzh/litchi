# 0749 log sections (ready to paste)

## For HOTSPOTS.md

## 0749 — Reuse-plan validation compares streams in place

[0749](0749-cfb-reuse-plan-validation.md) removes the copying half of the
CFB Reuse-plan validation that 0733 put at 51% of the PPT finish's CFB write.
`ReusePlan::validate` keeps its structural reparse of the planned view. Its
stream readback no longer materializes each stream through `open_stream`. A
`StreamComparer` bound to the reparsed view performs `open_stream`'s
traversal, with its checks and errors, loading the root mini stream once as
`open_stream` does. It compares each physical range where it lies, and a
correctly placed payload run compares by identity.

The first version was quadratic in mini streams (review: 3,000 mini streams,
8.2 → 114 ms); the review round made the readback linear (×3.07 Callgrind
instructions from 1,000 to 3,000 streams). On 45543.ppt, validation fell
from 732.5 K to 163.4 K Ir per write.

What remains:

- The unchanged reparse's per-stream table-sized map clearing (A5,
  `prepare_visited`) is now the one super-linear term, and it runs on every
  `OleFile::open`.
- Planning is the largest phase of the CFB-only write.
- The common editor still grows its output `Vec` from empty.

OLE2 lifecycle timings on this host move with glibc's heap-trim and mmap
thresholds; fixing both at 256 MiB removes the page-fault differences. The
non-iWork goal remains active.

## For REPORT.md

## 0749 — in-place Reuse-plan validation retained after a review fix

[0749](0749-cfb-reuse-plan-validation.md) keeps the CFB Reuse-plan
validation's proof and removes its copies. The structural reparse is
unchanged, and each stream is compared in place instead of read into a
fresh buffer.

The review found the first version quadratic in mini streams (v3, 3,000 ×
2,000 B: 8.2 → 114 ms; all-mini-stream files median +7.4%). The fix loads
the root mini stream once per validation and compares contiguous mini
sectors as one range. The re-measurement ran three arms (base, first
version, fix), 12 paired layouts on core 20 and 660 processes, and output
bytes were identical throughout.

| Reuse `write_to` | Fix vs base, Δ p50 |
|---|---:|
| 45543.ppt, 41246-1.ppt, FloatingPictures.doc, NoHeadFoot.doc | −40.9%, −33.5%, −32.1%, −16.9% |
| hyperlink.doc, empty.ppt, WithCheckBoxes.xls (two edits each) | −3.9% to −8.2% |
| generated v3/v4, 1,000 to 10,000 mini streams | −6.4% to −26.1% |

Over the 214-fixture corpus, both edits, the fix runs a median −12.5% of the
base's instructions; the worst of 418 pairs is +1.22%. The first version was
+14.1% on all-mini-stream files.

The fix also closes a pre-existing gap: a planned view that exceeds the
plan's own limits (`LimitExceeded`) now declines to the from-scratch writer
instead of failing the save.

Flags:

- the harness DOC `large` save, from glibc's trim policy (equal
  instructions; none with fixed thresholds);
- reader and Rewrite controls, from code placement (equal instructions).

Gates pass for `litchi-cfb` and its OLE2 dependents. `performance_claim:
none`. [Evidence](results/change-0749/README.md).

## For GOAL_AUDIT.md

## 0749 — CFB validation preserved, readback made copy-free and linear

[0749](0749-cfb-reuse-plan-validation.md) states what the Reuse-plan
validation proves: the structural reparse (header, FAT, directory tree,
MiniFAT, exact acyclic non-overlapping chains, physical partition) and the
readback (each model path is a stream of the model's length and bytes, with
every touched range readable). It also states which later obligations
depend on each part.

The reparse is unchanged. The readback reaches the same verdict without
copying, call for call with `open_stream`, including its one-time
mini-stream load. The old validator is kept as a test oracle. So every CFB
validation, ownership, cycle, overlap, FAT, MiniFAT, directory and
truncation check the goal requires is preserved (ADR 0006, ADR 0005's
validated-before-sink rule, ADR 0026). Test-only fault injection (5,600
random plan faults and named faults) shows each harmful fault still
refused.

The review's quadratic mini-stream cost, an adversarial CPU cost as well as
a regression, is fixed and tested for work per stream. A plan-derived
`LimitExceeded` now declines as the documented contract says.

Remaining debt:

- the reparse's A5 per-stream map clearing, which predates this record and
  sits on every reader open;
- placement-only read-control flags;
- planning, now the largest phase of the Reuse write.

No coverage, claim or timing-contract promotion; the non-iWork goal remains
open.
