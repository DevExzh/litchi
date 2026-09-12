# 0543: XLSX shared traversal cap-boundary rejection

The measured candidate is rejected. Original pilot gates passed: planning p50
fell 28.5–31.6%, workflow p50 fell 5.9–11.9%, and planning instructions fell
24.9–25.4%. However, the supplemental valid-input cap sweep found planning p50
regressions of 59.0–61.4% just above the 131,072-event provisional cap and
23.9–25.8% at 256×256 cells. Warning-denied Clippy also rejected the private
large parse-result enum. No runtime speedup is retained.

Production is restored to the baseline. A public post-EOF raw-error regression
test remains, covering both main worksheet shapes, exact typed refusal, retry,
source identity, unaffected-sheet recovery and byte exact no-op publication.
The temporary cap example is removed from live source; its source and measured
binaries' identities remain in this evidence bundle.

- [Decision](decision.json), [quality](quality-summary.json)
- [Frozen main plan](plan.json), [protocol](protocol.md)
- [Native](comparison.json), [allocation](allocation-analysis.json), [refusal](guard-analysis.json)
- [Planning profile](planning-profile-analysis.json), [profile review](profile-review.md)
- [Hardware](hardware-analysis.json), [eager controls](eager-analysis.json)
- [Cap plan](cap-boundary/plan.json), [cap results](cap-boundary/cap-analysis.json)
- [Individual reviews](adverse-review.json), [hardware review](hardware-review.json), [cap static review](cap-boundary/review.md)
- [Next priority](next-priority.md)

The invalid-input gates use the frozen baseline-valid envelope. Passing them
must not be read as unchanged invalid-input cost: late-validator p50 grows
174.5–184.7%, and allocated bytes grow about 616–961% against the same invalid
baseline. Those costs are explicitly reviewed. Normal and allocator timing
vectors are separate; instrumentation time is not a native latency claim.

Three failed command receipts are preserved: an initial public-test fixture/API
mismatch, candidate final Clippy, and the cap harness's initial compile failure.
Corrected source freezes precede the respective successful measurements. No
failed or superseded evidence is erased. The main allocation status-string and
cap analyzer corrections preserve their original analyzer bytes and provenance;
they do not change numerical admission thresholds.

Run `python3 -B verify.py --strict` to replay sealed evidence after cleanup.
Historical capture drivers refuse to overwrite receipts. Precleanup/preseal
checkpoints intentionally record their then-pending cleanup/seal states; the
final seal and strict replay establish final custody. Follow-up proposals are
unmeasured and require a new source-bound campaign.

OLE2/OOXML optimization remains active; ODF is deferred until that goal completes.
