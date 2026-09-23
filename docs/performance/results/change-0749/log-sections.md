# 0749 log sections (ready to paste)

## For HOTSPOTS.md

## 0749 — Reuse-plan validation compares streams in place

[0749](0749-cfb-reuse-plan-validation.md) removes the copying half of the
CFB Reuse-plan validation that 0733 put at 51% of the PPT finish's CFB write.
`ReusePlan::validate` keeps its structural reparse of the planned view. Its
stream readback no longer materializes each stream through `open_stream` (a
zero-filled buffer, a per-sector clear and copy, and a whole-stream compare
per stream). `OleFile::stream_equals` performs `open_stream`'s traversal,
with its checks and errors, and hands each physical range to the plan, which
compares it where it lies. A correctly placed payload run is the model's own
slice and compares by identity.

On 45543.ppt, validation falls from 732.5 K to 163.4 K Ir per write, and the
CFB-only Reuse `write_to` from 31.9 to 22.0 µs. Planning is now the largest
phase of that write, about 45% of native samples. Its main costs are the
`place` closure's per-sector checks, per-element chain growth,
`high_water` scans and path-keyed maps, and they are the next candidate. The
common editor's `render_with_layout` still grows its output `Vec` from empty.

OLE2 lifecycle timings on this host move with glibc's heap-trim and mmap
thresholds. Fixing both at 256 MiB removed a 478–868 page-fault-per-owner
difference in `doc_semantic_one_edit_save/large`; 0745's inferred mechanism
is now tested. The non-iWork goal remains active.

## For REPORT.md

## 0749 — in-place Reuse-plan validation retained

[0749](0749-cfb-reuse-plan-validation.md) keeps the CFB Reuse-plan
validation's proof and removes its copies: the structural reparse is
unchanged, and each stream is compared in place instead of read into a fresh
buffer. Output bytes are identical in every measured process (492 in the
main matrix, 64 in the fixed-threshold check). Over 8 paired heap layouts on
core 20:

| Lane | Fixture | Paired Δ p50 |
|---|---|---:|
| CFB-only Reuse `write_to` | 45543.ppt, 41246-1.ppt, FloatingPictures.doc, NoHeadFoot.doc | −30.7%, −33.5%, −23.5%, −17.1% |
| 0728 container control, Reuse | the three 0728 fixtures | −15.2%, −7.3%, −6.1% |

Per write on 45543.ppt, 137 K fewer user instructions and 386,906 fewer
allocated bytes.

The public lifecycles perform one such write each. Over 18 layouts, the PPT
removal changes −0.7% [−2.2, +3.2] on 45543.ppt and −1.05% [−1.24, −0.71] on
41246-1.ppt; the DOC replace changes +0.05% on FloatingPictures.doc and −6.1%
on NoHeadFoot.doc.

Flags:

- The 45543.ppt removal's p95 is +6.96%.
- The harness `cfb_open/few-large` read control is +4.7% with identical
  instructions. It is code placement: 35% more L1i fills and 49% more branch
  misses.
- `doc_semantic_one_edit_save/large` shows up to +21% in some layouts. That
  is glibc's trim policy (tested with fixed thresholds), not work.

With `glibc.malloc.trim_threshold` and `mmap_threshold` fixed at 256 MiB, the
45543.ppt removal changes −3.50% [−4.19, −2.91] with no p95 flag, and the
harness DOC `large` save −1.96%.

The verdict equals the old readback's on 5,000 random plan faults, named
overlap, cycle, length, content, mini-stream, header and table faults,
9,041 corpus stream comparisons and 2,884 corrupted files. Gates pass for
`litchi-cfb`, its eight dependents and the facade. `performance_claim:
none`. [Evidence](results/change-0749/README.md).

## For GOAL_AUDIT.md

## 0749 — CFB validation preserved, readback made copy-free

[0749](0749-cfb-reuse-plan-validation.md) states what the Reuse-plan
validation proves: the structural reparse (header, FAT, directory tree,
MiniFAT, exact acyclic non-overlapping chains, physical partition) and the
readback (each model path is a stream of the model's length and bytes, with
every touched range readable). It also states which later obligations depend
on each part.

The reparse is unchanged. The readback reaches the same verdict without
copying, and the old validator is kept as a test oracle. The new path has no
public API, dependency, `unsafe` or limit change, and emitted bytes are
unchanged, so every CFB validation, ownership, cycle, overlap, FAT, MiniFAT,
directory and truncation check the goal requires is preserved (ADR 0006,
ADR 0005's validated-before-sink rule, ADR 0026). A test-only fault
injection over real plans shows each harmful fault still refused.

Remaining debt:

- a code-placement read-control flag (+4.7%);
- the 45543.ppt removal's p95 flag;
- planning, now the largest phase of the Reuse write.

No coverage, claim or timing-contract promotion; the non-iWork goal remains
open.
