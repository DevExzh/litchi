# Ready-to-paste log sections for change 0763

## HOTSPOTS.md

## 0763 — per-charge budget atomics in the streaming DOCX and XLSX writers

[0763](0763-rough-budget-leases.md) applies owner decision 6 of 0758: the two
writers charge Objects, Work and (DOCX) input bytes through rough budget
leases that claim 4 Ki / 64 Ki / 64 KiB chunks and hand units out locally.
A large DOCX iteration makes 299 budget atomics instead of about 1.05
million, a large XLSX one 307 instead of about 787,000. The gain is smaller
than the flat profiles suggested: about 8 cycles per removed atomic, −6.0%
on large DOCX creation and −1.2% on XLSX. Disproven: `consume`'s 15–25% share
of the flat DOCX profile was not its own cost; a locked update also absorbs
stalls on the stores before it. What remains in the streaming writers is
Deflate (0762), the stage and escaping copies, and PPTX's per-member costs.

## REPORT.md

## 0763 — rough budget leases

[0763](0763-rough-budget-leases.md) is retained with `performance_claim: none`.
Base `6ec785c265` (record 0762's tip); commit `1ab98ca669`. With both legs
built by the same command, `docx_streaming_create` moves 36.501 → 34.217 ms
(−6.03%, CI [−7.15%, −5.79%]) large and 2.345 → 2.174 ms (−7.51%) medium;
`xlsx_streaming_create` 90.976 → 89.928 ms (−1.24%) and 5.704 → 5.620 ms
(−1.30%). PPTX (no budget charges) and the preservation controls stay within
noise; no allocation changes; every output byte is unchanged. Instructions
fall 0.3–0.6%; the atomics' cost was in cycles. A base-versus-candidate probe
over 28,860 limit and cancellation scenarios prints byte-identical
transcripts.

## GOAL_AUDIT.md

## 0763 — chunked budget charging with exact limits

[0763](0763-rough-budget-leases.md) adds `Budget::lease` and
`ExecutionContext::lease` under owner decision 6 of 0758 and a dated
clarification to ADR 0005. A claim is the same atomic check-and-add per level
as `consume`, shrinks to the room left but never below the charge's need, and
rolls back on refusal. No level is ever over its limit (a monitor thread
checks this under eight concurrent threads mixing leases and reservations).
The holder's refusals are those of exact accounting, and unspent units return
on release, drop, error, cancellation and poison. A 28,860-scenario differential
against the base proves the sole-holder behaviour identical. Other holders of a
shared budget see up to one chunk per writer and resource pre-claimed, as
decision 6 accepts. Cancellation is still checked before every charge.
