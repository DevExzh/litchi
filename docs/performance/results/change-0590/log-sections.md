# Log paragraphs for change 0590

The coordinator merges these into the four logs; this batch did not edit them.
Each is written in the style of the newest section of its target file.

## For `docs/performance/HOTSPOTS.md`

## 0590 — PPTX opened-transaction revision reuse

Item PPTX-1(a)/(b) of the 0587 queue is implemented. An opened PPTX lifecycle
hashed the complete package — every part blob, media included — six times and
captured the package five times; it now hashes four times and captures four.
`Transaction::commit` keeps the revision it computed for the `unsign()`
decision and reuses it for the recapture, recomputing only when `is_signed()`
says stripping may have rewritten a fingerprint input, and
`Package::apply_opened_presentation_commit` stops discarding `commit.snapshot`:
it still builds, validates and assigns exactly the same candidate, then reuses
that already-validated capture when a direct comparison of every
`package_fingerprint` input proves the candidate byte-identical, and captures
from scratch otherwise. **Measured** on the isolation pair
(`--samples 3` minus `--samples 1`, callgrind, CPU 10): per lifecycle of
`pptx_eager_batch_edit_save` 12,679,371,772 → 10,680,374,116 Ir, −15.77%;
whole child 39.61 G → 33.61 G (−15.15%), `package_fingerprint` inclusive 34.94%
→ 24.76% and software SHA-256 51.77% → 44.60%; `pptx_eager_multi_slide_batch_edit_save`
−15.12%, `pptx_slide_remove_boundary_save` −1.80%,
`pptx_slide_move_boundary_save` −2.57%. The revision value, the durable
`LPRM0001`/`LPCP0002` patch encodings and every refusal are unchanged; 860
`litchi-pptx` tests pass, six of them new. The eager corpora are built without
physical source provenance and re-deflate all media on save, so the timed region
is dominated by work this change does not touch and the wall-clock deltas sit
inside the host's A/A floor. Item (c) changes the durable revision format and
stays a frozen-design prerequisite; item (d) was examined and rejected.
`performance_claim: none`. [Change and limitations](0590-pptx-opened-transaction-revision-reuse.md);
[evidence](results/change-0590/README.md). OLE2/OOXML remain active; ODF is
deferred until completion and iWork excluded.

## For `docs/performance/GOAL_AUDIT.md`

## 0590 — PPTX opened-transaction revision reuse

`docs/GOAL.md`'s first optimization step — eliminate unnecessary work — applied
to the PPTX opened-document CRUD route that 0587 ranked third. Two of the four
complete-package SHA-256 passes per lifecycle are removed by reusing values
already computed on identical content: the commit's own unsign-check revision,
and the commit's own capture, which the facade had been discarding in favour of
re-deriving it. Both reuses are conditional and fall back to the original path,
so no refusal, no output byte, no limit and no validation moved; ADR 0003's
complete-package revision binding is preserved because the fingerprint
definition is untouched and the durable patch encodings that carry revisions are
bit-identical. **Measured** −15.77% Ir per lifecycle on `pptx_eager_batch_edit_save`
with exact before/after call counts (six hashes and five captures per lifecycle
before, four and four after). What stays open: the remaining four hashes need
0587's item PPTX-1(c), which redefines the durable revision and therefore needs
a frozen design record and a magic bump for `LPRM0001` and `LPCP0002` before any
code; the notes-index memo of item (d) needs an owner for the memo and an
ADR 0013 invalidation rule; the cross-package copy path (PPTX-2) is untouched;
and no eager PPTX save timing is believable until the `open`/`from_vec` corpus
variant 0587 asks for exists, because the present corpora re-deflate all media.
No speedup, RSS, allocation, cold-cache or real-producer claim follows.

## For `docs/performance/REPORT.md`

## Change 0590: PPTX opened-transaction revision reuse

Change 0590 removes two of the four complete-package hashes in the
`litchi-pptx` opened-presentation lifecycle. `Transaction::commit`
(`opened/transaction.rs`) retains the revision it computes to decide
`unsign()` and passes it to the recapture through a new
`capture_with_revision`, recomputing only when `OpcPackage::is_signed` reports
that stripping may have rewritten a fingerprint input; every structural, limit
and notes-topology check in the capture still runs unconditionally, and a
`debug_assert!` re-hashes in debug builds. `Package::apply_opened_presentation_commit`
(`package/model.rs`) now forwards `commit.snapshot` to the publication helper,
which reuses that capture only when a new `packages_equal` — a direct comparison
of exactly the inputs `package_fingerprint` feeds, short-circuiting on blob
`Arc` identity — proves the built candidate byte-identical, and otherwise
captures it exactly as before. Validation passed `litchi-pptx` library
`556/556` plus integration and doc tests, 860 in total, with six new tests
covering a signature-stripping commit, the revision binding of an ordinary
commit, snapshot equality with a fresh capture on both the changed and no-op
routes, drift outside the patch write set, a stale write set, and the
exhaustiveness of the package comparison. **Measured**, callgrind isolation pair
on CPU 10: −15.77% Ir per `pptx_eager_batch_edit_save` lifecycle, with the six
hashes and five captures per lifecycle becoming four and four. The paired native
timings are reported beside this host's A/A floor and are not a speedup claim;
the eager corpora re-deflate all media on save. See
[Change 0590](0590-pptx-opened-transaction-revision-reuse.md);
`performance_claim: none`.

## For `docs/performance/ADR_COMPLIANCE.md`

## Change 0590 compliance update

Change 0590 keeps ADR 0003's revision binding exactly: the complete-package
revision is still `package_fingerprint` over the same domain string, the same
sorted part feed of name, content type, payload and relationships, the same
package-root relationships and the same opaque non-part members. No revision
value changes, so the durable `SlideRemovalPatch` (`LPRM0001`) and
`CrossSlideCopyPatch` (`LPCP0002`) encodings that carry revisions remain
bit-identical and patches serialized before the change still apply after it.
ADR 0003 specifies that `commit()` returns a `Commit<T>` "containing the new
snapshot, a reversible patch, and diagnostics"; part (b) of this change makes
the PPTX facade use that snapshot rather than discard it, and reuses it only
after proving the published candidate byte-identical to the package it
describes. ADR 0005's mandatory validation is intact: no check is removed,
weakened, reordered or made conditional, and both reuses fall back to the
original path on any mismatch, including a mismatch in resource limits or the
physical-source-provenance flag. ADR 0006 preservation is intact because no
output byte changes and the `unsign` signature policy still forces a fresh hash
whenever stripping rewrites bytes. ADR 0013's notes-topology check still runs
inside every capture, including the capture that is later reused; the change
declines to capture the same content twice, never to capture it zero times. No
`unsafe`, no new ambient I/O, no global pool, no weakened limit, and no public
leakage of archive types, locks or executors. See
[Change 0590](0590-pptx-opened-transaction-revision-reuse.md);
`performance_claim: none`.
