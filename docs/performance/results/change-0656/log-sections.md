# Log paragraphs for change 0656

Four paragraphs, one each for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and
`ADR_COMPLIANCE.md`, in the style of their newest sections. The coordinator
merges these; this change does not edit those files.

## For `HOTSPOTS.md`

**Queue row 5 is closed, and the cross-package copy's hotspot has moved to
planning.** 0646 priced the second candidate build and this change removes it:
`bounded_package_bytes` falls from two calls per lifecycle to one,
`zlib_rs::deflate::deflate` from 1,112 calls to 556 on the media-rich corpus and
from 40 to 20 on the plain one, and every other call count on the path — eight
`package_fingerprint`, four `physical_package_fingerprint`, eight
`capture_internal`, two `build_candidate`, two `from_vec_reusing_payloads` — is
identical in both legs. Per lifecycle that is −11.44% of instructions on the
media-rich corpus and −2.42% on the plain one; natively it is −24.6% to −27.1%
of whole-child cycles against A/A floors of 0.8–1.7% in three independent
windows, and the phase it lands in, apply, falls **70.55–73.41%**
(328.2 → 89.4/90.6 ms p50) across four. The `Ir`-versus-cycles gap is SHA-NI, exactly as 0646 predicted:
valgrind masks the CPUID bit, so every instruction share this area's records
attribute to hashing is about five times its native cycle share. **Rank this
path off `perf stat`, not off the `Ir` tables.** What the second serialization
was hiding is now the top term: **planning is the largest remaining native cost
on a media-rich lifecycle**, a p50 of about 297 ms against apply's 94 ms, and it
still builds, deflates and reopens a whole candidate — with no way to avoid it,
because the plan's whole purpose is to prove that candidate exists. The next
item in this area remains 0587's PPTX-1(c), the per-part digests with a redefined
complete-package revision, which change 0655 lands in this wave. One new
observation for the queue: `publication_ns` on the media-rich selectors is
**bimodal at ≈1.5 ms and ≈6.4 ms independently of which binary runs it**, and
this change has two within-window witnesses — one leg of each binary in each
mode inside one run, and an after *pair* split across the two modes in
another. It executes no code either 0646 or 0656
changed and it is unexplained; any later record that reads that phase should
expect the ±4× and not attribute it.

## For `GOAL_AUDIT.md`

Change 0656 is the answer to the gate change 0646 could not clear, and the shape
of the answer is worth recording. 0646 stopped at "an optimization that changes
how long memory lives must be opt-in until something is charging for it." The
owner's decision 5 supplied the charge, and the charge turned out not to be a
budget at all: it is a **declared ceiling in the operation's own finite limit
policy**, `opened::Limits::max_retained_candidate_bytes`, the sibling of the
`max_history_bytes` that `History::push` already enforces. The crate now has two
enforced live-memory ceilings instead of one, both in `Limits`, both intersected
by `intersect_limits`, and both reached by no `Budget` — which is exactly the
gap ADR 0005 names and exactly the interim location the amendment now writes
down until ADR 0031's execution context reaches this path.

Two rules came out of the implementation that generalize past this change.
**First: a retention ceiling whose alternative is recomputation must not
refuse.** 0646 specified a `Require` route with a typed `Error::Limit`; decision
5 removed it, and it was right to. A typed limit error protects a caller from
work or memory it did not ask for, and there is nothing to protect it from when
the fallback is the path the library takes anyway. The over-budget case
therefore changes no result, no refusal and no published byte, and it is
observable through `retained_candidate_bytes()` rather than through an error.
**Second: derived state must not be part of a value's identity.** 0646's scratch
compared a stored digest in `PartialEq`, which made a plan that had released its
archive unequal to the plan it was; this implementation's equality ignores the
slot entirely, so `release_retained_candidate` is value-preserving and
`plan_a == plan_b` can never become a byte comparison of two archives. The same
discipline caught a second trap: a derived `Debug` would have printed 33.6 MB of
archive, so `Debug` reports the retained length instead. **The measured cost is
disclosed rather than netted out**: a media-rich plan-plus-apply lifecycle now
peaks 12.18% higher, a plain one peaks 5.45% lower, and the budget is the knob
that buys the old profile back.

## For `REPORT.md`

**0656 — the cross-package copy stops deflating the candidate twice, under a
64 MiB retained-candidate budget.** `CrossSlideCopyPlan` keeps the serialized
candidate archive that planning already built — the same allocation the
candidate reopen holds, taken as a second owner through a new
`OpcPackage::exact_source_shared`, so retention is a decision not to free rather
than a copy — and `apply_cross_slide_copy_plan` reuses those bytes instead of
serializing and deflating the candidate again. The hold is bounded by a sixth
member of `opened::Limits`, `max_retained_candidate_bytes` (default 64 MiB),
intersected between the two snapshots; a candidate above it is simply not
retained and its application rebuilds, which is never a refusal. Measured on the
0646 selectors, pinned to CPU 11, with the A/A floor of the same window:
`pptx_cross_copy_media_rich` 663.198 → 417.352 ms p50 (**−37.07%**, A/A −2.36%),
`…_lifecycle` 673.748 → 413.485 ms (**−38.63%**, A/A +4.71%),
`pptx_cross_copy_plain` 8.048 → 7.676 ms (−4.62%, A/A −1.35%), `…_lifecycle`
9.248 → 9.102 ms (−1.58%, A/A −3.12%), reproduced in a tighter-floored window
(A/A p50 under 2% everywhere) at −35.04%, −33.38%, −2.91% and −2.00%; the apply
phase alone falls 70.55–73.41% across four windows and whole-child `perf stat`
cycles fall 24.64%, 27.01% and 27.08% in three. Deterministic counts: deflate calls per lifecycle
1,112 → 556 media-rich and 40 → 20 plain, one serialization removed and nothing
else moved. The cost, measured to the byte: the plan holds the archive plus a
40-byte `Arc` header (33,599,873 + 40 on the media-rich corpus, 31,545 + 40 on
the plain one) and releases it exactly on drop; the media-rich lifecycle's peak
live bytes rise 12.18% while the plain one's fall 5.45%, and total allocated
bytes fall 19.98% and 15.44%. **Breaking:** `Limits::new` takes six arguments
and `Limits::DEFAULT` changes value. Published bytes are identical on every
route (one `output_sha256` per corpus across sixteen timing legs, twelve `perf
stat` legs and twelve callgrind profiles); `CrossSlideCopyPatch` is untouched,
including whichever magic it carries after change 0655's independent bump;
`litchi-pptx` goes 875 → 885 tests with ten added and none changed. `performance_claim: none`.

## For `ADR_COMPLIANCE.md`

**ADR 0005 (I/O, memory, measured performance) — amended by this change, under
0652 decision 5.** The amendment is narrow and dated 2026-09-16. It says that
"cache behavior is semantically invisible" does not license retained state whose
size is the point; that such state is declared, bounded, observable and
releasable, with its ceiling in the operation's own finite limit policy where the
path has no execution context, and that those ceilings move to the execution
context when ADR 0031 reaches the path; and that exceeding such a ceiling falls
back to recomputation rather than refusing, because the typed `Limit` error has
nothing to protect a caller from when recomputation is always available.
Ceilings whose alternative is not recomputation keep their typed refusal. The
scratch-storage clause is untouched and is the reason the fallback is
recomputation rather than a spill: a candidate archive is a complete
presentation package, and a source ratchet test now refuses fourteen filesystem,
temporary-file, scratch-provider and mmap markers in `cross_copy_plan.rs`. The
limit-error clause ("resource, observed value, limit, and object path") is
unaffected here because this change adds no limit error; `litchi_pptx::Error::Limit`
still carries no observed value, which 0646 named as a pre-existing gap and this
change does not close.

**ADR 0003 (snapshots, edits, patches).** No revision value, proof format or
durable encoding changes. `CrossSlideCopyPatch`, its magic and its six 32-byte
revisions are untouched by *this* change and `apply_patch` retains nothing,
pinned by a test that compares the encoded patch byte for byte with and without
retention; change 0655 bumps that magic in the same wave for reasons unrelated
to retention, and the test checks the `LPCP` family prefix so it survives the
bump. The plan
remains an immutable value: the new field is private and plan equality ignores
it, so releasing the archive preserves the plan's value.

**ADR 0011 (OOXML physical package ownership) and the 2026-08-21 OPC
exact-source amendment.** The retained archive is the exact source the candidate
reopen already authorized, and `OpcPackage::exact_source_shared` returns that
authorization's own handle without changing what authorizes it; a later
revocation does not retract a handle already taken, which the accessor's
documentation states. Nothing about preservation provenance or the "planning
evidence only" rule changes. The one place this brushes ADR 0011 is the copy at
application, which remains: eliminating it needs a shared-bytes ingress
(`OpcPackage::from_shared_vec`), an ownership decision this change does not take.

**ADR 0001 and ADR 0005's no-leakage rule.** No archive type, raw lock or
executor enters a public API. `exact_source_shared` returns `Arc<Vec<u8>>` —
shared immutable bytes, not an archive type — and the plan exposes the hold as a
`usize` and nothing else.

**`docs/GOAL.md`.** Optimization-order step 2 (unnecessary I/O, decompression,
recompression). No new `unsafe`, no weakened limit or malformed-input defence,
no hidden global pool, no ambient I/O. Proposed ADRs are not cited as authority;
ADR 0031 is named only as where the ceiling goes when it arrives.
