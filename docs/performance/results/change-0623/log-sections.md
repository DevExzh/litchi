# Log paragraphs for change 0623

Four paragraphs, one each for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and
`ADR_COMPLIANCE.md`, in the style of their newest sections. The coordinator
merges these; this change does not edit those files.

## HOTSPOTS.md

**ZIP-5 is closed, and with it change 0577.** The OOXML open must read every
relationship part — 0577 settled that, against ADR 0005 and ADR 0006, and its
`e4`/`e5` pair showed the admission verdict has no later point to move to. What
0577 could not do was reduce their *shape*, because the primitive it needed did
not exist: the read path held no way to ask where a member's local record is.
`IndexedArchive::local_span_hint` is that primitive — one offset and one length,
facts the index already holds and no ZIP framing — and `litchi-opc` now uses it
to fetch each physically contiguous run of structural members in one read before
a walk that will revisit them in an order of its own. Measured on change 0587's
own probe: a source-backed open of the 132-member workbook falls from **45
requests to 10** and `shapes.pptx` from **24 to 6**, both for **exactly the same
bytes** — the run read is the union of the per-member spans change 0611 already
reads, so no byte is fetched that was not fetched before. 0577 modelled 10 on
the workbook and 10 is what it costs. Across all 533 ZIP containers under
`test-data` the open's requests fall 4,160 → 2,569 (−38.2%; −45.2% over the 321
OOXML-extension containers) with **zero** containers costing more requests, and
the corpus reads 1,522 bytes *fewer* in total. `comment.docx` does not move: its
three structural members sit in three separate runs, which is the degenerate
case 0577's admission gates required be shown to cost no more. The remaining ZIP
item on the 0587 queue is **ZIP-6**, the tail-window locate, whose three
locator reads are now 30% of what a workbook open costs.

## GOAL_AUDIT.md

Change 0623 is the first change in this programme to land a frozen design
written by an earlier record without amending it, and the reason is worth
recording: 0577 froze not just the mechanism but the *invariants*, and two of
them turned out to do all the work. Invariant 1 — "every byte read lies inside a
structural member's own local record" — was written before change 0611 existed,
and after 0611 it has a sharper form that can be *computed* rather than argued:
join two members into a run only when the earlier member's own first-read span
already reaches the later member's local header. That single rule makes the run
the union of spans the archive already reads, which is why the measured byte
totals are identical rather than 15% higher as 0577 modelled, and why no
container in the corpus reads a byte the before leg did not. Invariant 5 — "a
run read that fails falls back to per-member reads" — is the invariant change
0611 explicitly *refused* for its own one-member read, and both records are
right: a read that is a member's own first contact with the source must not be
retried, and a read that belongs to no member must be abandonable. A design
record that states its invariants precisely enough to be contradicted by a later
change is more useful than one that states a mechanism.

The second lesson is about where an optimization is allowed to apply. This one
is gated to the ordinary unmanaged, exact-policy open. A speculative read
spanning several members is reserved and committed against
`Resource::InputBytes` and observed for cancellation as a unit, so under an
`ExecutionContext` it would move where those observations land — and change
0611's justification for accepting that class of movement (its span read *was*
the member read, one for one) is not available here. Declining to apply a saving
on the managed path is a cheaper answer than sixteen corrected test triggers,
and it is the honest one: the mechanism genuinely is not invisible there.

## REPORT.md

**0623 — one read per contiguous run of structural members.** Retained,
implemented in `soapberry-zip` (`office.rs`, additive only: `LocalSpanHint` and
`IndexedArchive::local_span_hint`) and `litchi-opc` (`source_backed.rs`,
`source_backed/read_ahead.rs`), `performance_claim: none`. This is change
0577's candidate (c), implemented as frozen, including its intra-run
error-precedence decision. Deterministic counts on change 0587's probe: open 45
→ 10 requests on the 132-member workbook and 24 → 6 on `shapes.pptx`, both at
**identical bytes**; `comment.docx` unchanged at 6, its runs being degenerate.
Across 533 ZIP containers: 4,160 → 2,569 open requests, 0 costing more, 16
reading fewer bytes, 2 reading more (+332 B and +594 B, both packages whose open
*refuses* part-way through a prefetched run), and source observations unchanged
on every one. Correctness: an open differential over all 533 containers,
comparing the open's verdict, every package and part relationship, every
admitted Part with its content type, every non-part member and the decoded
length and CRC of every Part — the two reports differ on **nothing but the
request count**. Paired ABBA timing on the 1 ms-per-request simulated transport
of changes 0493 and 0572, 30 samples per leg, CPU 17, A/A floor under 0.1% at
p50. Nineteen new tests in `litchi-opc` and three in `soapberry-zip`; no existing
test was changed. Gates: thirteen sections, all exit 0.

## ADR_COMPLIANCE.md

**ADR 0005 (bounded resources, caching semantically invisible).** Two named
ceilings, both measured rather than guessed against the 533 containers under
`test-data`: `MAX_STRUCTURAL_PREFETCH_RUN_BYTES` = 64 KiB against a longest
observed run of 9,298 bytes, and `MAX_STRUCTURAL_PREFETCH_BYTES` = 256 KiB
against a largest observed retained set of 14,755 bytes. Every buffer is
reserved with `try_reserve_exact` and released when the catalog call returns, on
every path. "Cache behaviour is semantically invisible" is measured, not
asserted: the open differential over all 533 containers reports identical
verdicts, relationships, part catalogs, non-part members and decoded payload
CRCs. **ADR 0006 (validation, fail-closed).** Every check keeps its position and
its identity; only the fetch is reordered, never the walk. The `e4`/`e5`
admission pair change 0577 made decisive is pinned as a regression test, as are
a malformed relationship part, a duplicate relationship ID and a relationship
budget refusal at the start, the middle and the end of a coalesced run, each
shown to reach the same verdict with the coalesced fetch and with it refused.
**ADR 0011 (ownership).** `soapberry-zip` stays the ZIP grammar owner. What
crosses the boundary is an offset and a length in the caller's own byte source —
where bytes are, not what they mean — which is the ownership-respecting addition
change 0577 said the design needed. The `soapberry-zip` diff is additive: no
existing line changes, so no read path in that crate can have moved. **Change
0317 (error precedence).** Unchanged: the prefetch is a separate read that
belongs to no member, and a failure of it is abandoned rather than reported, so
every member read keeps exactly one refusal, one cancellation observation and
one resource reservation. **One stated scope limit.** The mechanism is off for
a managed open and for an explicitly configured forward window, because a
multi-member speculative read is budgeted and observed as a unit and would move
where a managed refusal fires.
