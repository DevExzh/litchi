# 0758 — the owner's decisions of 2026-09-24: ADR 0032 accepted, 0751's reading confirmed, gate G5 kept, three new policies authorized

Status: retained, decision record. `performance_claim: none`. This record
carries no measurement. It records the owner's decisions and is the authority
that the records implementing them cite.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What this record is

Integration record [0756](0756-ole2-ooxml-wave-integration.md) closed the
0742–0757 wave with six questions only the owner could answer. On 2026-09-24
the owner answered all six and asked for the next wave to proceed and for
branch `feat/spec-gap-implementation` to be merged. The record states:

- how the programme reads each answer;
- what it authorizes;
- what an implementing record must still prove.

With this record, proposed ADR 0032 becomes Accepted and ADR 0005's
2026-09-16 memo amendment gains the confirmed reading. The standing trade-offs
of [0652](0652-owner-decisions-for-the-third-wave.md) still bind every decision
below: breaking changes are acceptable, correctness and safety come first, and
most inputs are benign.

## The decisions

| # | The owner's words | 0756 question | What it authorizes | What the implementing record must still prove |
| ---: | --- | --- | --- | --- |
| 1 | "Accept ADR 0032." | 1; [0743](0743-pptx-semantic-text-and-edit-path.md) | [ADR 0032](../adr/0032-snapshot-derived-value-memos.md) is Accepted as of this record. Snapshot memos may hold small derived values of payload bytes (not only digests) under the memo amendment's conditions. Re-applying 0743's withdrawn slide-root memo (`99ce9c5e34`) is authorized. | Every ADR 0032 condition, proven rather than asserted. Reservation failure is a typed resource error, not a silent empty memo. No refusal, value or published byte depends on memo contents. The one-edit commit effect is re-measured on the current tip. |
| 2 | "Confirm the interpretation of 0751." | 2; [0751](0751-pptx-cross-copy-apply-digest-reuse.md) | A facade may fill an empty memo slot from a capture of its current graph, provided every kept entry is re-projected onto the facade's own allocations. The reading is written into ADR 0005's 2026-09-16 memo amendment as a clarification. | Nothing new: 0751 already proves re-projection at every adoption point. Later memo adopters cite this clarification. |
| 3 | "Reject relaxing the Gate G5." | 3; [0646](0646-pptx-cross-copy-candidate-retention-design.md) | Nothing. The retained cross-copy archive's digest stays recomputed at application, over the bytes about to be published. | — |
| 4 | "Accept the batch compression policy." | 4; [0752](0752-streaming-writer-small-write-batching.md) | Creation-from-scratch writers (streaming DOCX/XLSX/PPTX and fresh DOC/PPT/XLS/OOXML writers) may produce new output bytes: coalescing small writes before the compressor, or cheaper per-member compression strategies. Preservation of existing packages is unaffected: untouched members, source-backed copies and regenerated members of opened packages keep today's rules. | Output stays a pure function of the input (independent of caller write sizes, thread count and hash seed), and is proven so across processes. Every changed output is read back semantically identical. Size changes per corpus are reported; owner decision 10 of 0652 prefers smaller files, so the record justifies any growth above 1%. Durable patch formats that bind physical output, if any are affected, bump or are shown unaffected. |
| 5 | "Accept the opt-in save policy." | 5; [0714](0714-docx-atomic-publication-attribution.md) | An explicit, opt-in durability policy on atomic filesystem saves that skips or weakens the file and directory syncs. The default stays fully durable, unchanged. The policy flows through an explicit caller-supplied option, not an ambient setting (ADR 0005). | The default is proven unchanged (same syscalls). Each level's sync behaviour is verified with a syscall trace or a counting test double. Atomic replacement (no torn file on crash before rename) is kept at every level. Errors stay typed. The record states what each level does and does not guarantee. |
| 6 | "Accept a rough budget lease rather than a precise one." | 6; [0752](0752-streaming-writer-small-write-batching.md) | Budget accounting may be chunked. An operation acquires a lease of up to a chunk of budget and consumes it locally. Siblings observe the pre-claimed amount. A refusal may come earlier than exact accounting would give it, at lease granularity. | A limit is never exceeded: a lease never claims more than remains, and consumption never exceeds the leased amount. Unused lease returns on drop, including on error and cancellation. `ResourceLimit` errors stay typed. The lease size is stated with its reason. Concurrency tests show every counter settles exactly. |

## The merge

The owner asked for the work on `feat/spec-gap-implementation` to be merged
into `feat/office-format-completeness`. Change
[0759](0759-spec-gap-branch-merge.md) records that merge. It covers the
committed tip `a67a38abf2` only; the uncommitted changes in that branch's
worktree belong to its agent and are left untouched.

## What this record does not decide

Everything else in 0756's follow-up queue proceeds under the standing rules,
without a new decision: the quick-xml attribute hash exposure, BOM-prefixed
parts, the XLS writer length-field bugs, and the remaining performance
follow-ups.
