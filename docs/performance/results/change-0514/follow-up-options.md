# 0514 follow-up options for the XLSX parser/layout fusion

`scope: bounded future decision record`

`base revision: 33a21e0f0; rejected candidate archived, production restored`

`performance_claim: none; candidate rejected at pilot admission`

This note evaluates the next action for the private shared-event worksheet
parser and snapshot `Layout` prototype. It reads the 0512 operation profile,
the 0513 allocation-boundary evidence, the 0514 design and implementation
reviews, and the OLE2/OOXML queue. It makes no Rust change and runs no build,
profile, or workload. ODF remains deferred until the full OLE2/OOXML goal is
complete; iWork remains excluded.

## Current decision state

The candidate has a sound high-level mechanism: for an ordinary borrowed
MCE-free worksheet, one namespace-aware `NsReader` feeds the semantic parser
first and the lossless snapshot scanner second. The transaction retains the
scanner result only until an effective ordinary grid rewrite, then reuses the
ephemeral layout. MCE-transformed input, merge-derived bytes, source-identity
mismatch, metadata edits, and prepopulated stores use existing fallback paths.

The final [admission review](admission-review.md) rejects this prototype.
Cold same-value p50 increases 64.43–84.64%, with repeated dense incremental
peak increases of 27.04% and 18.64%. The 966-test owner suite and all-features
Clippy pass, but correctness does not justify the added speculative work.
Production has been restored. The conditional options below explain why the
current gate stops and what evidence a materially different follow-up needs.

The decisive gates are:

* the focused semantic and differential suite remains green, including exact
  output, source identity, MCE/fallback, event-limit, and typed error
  precedence;
* cold same-value cell/row/column actions do not pay an unacceptable
  speculative scanner/layout cost;
* the candidate's operation-local incremental peak,
  `region_peak_live_bytes - live_bytes_before`, has a practical bound against
  the separate-pass control; and
* changed-edit latency and allocation evidence show useful work elimination
  after every adverse row, including no-op and memory rows, is retained.

No branch below treats the candidate as accepted before those gates pass.

## Option A: fusion passes the cold no-op and memory gates

The next largest practical work is the **changed worksheet validation parse**,
measured at 25.77% of direct commit instructions in each 0512 repeat. Changed
XML compaction is the adjacent 20.38% child. These are disjoint direct-child
shares; nested parser, scanner, and compaction symbols must not be added to
them. The source/store and rewrite shares will change after fusion, so a fresh
operation profile should confirm the ranking before code selection.

The next investigation should therefore isolate the post-write path:

1. compact the changed worksheet while recording its own operation interval;
2. parse and style-validate the compacted bytes in a separate interval; and
3. retain publication/web proof and post-write semantic/readback checks as
   separately scoped controls.

Use the 0513 allocator region and the available hardware counter group inside
these operation boundaries. Keep absolute live/high-water values separate from
the checked incremental peak. The profile should determine whether the
compactor and changed-output parser repeat work that can share an
authenticated event interpretation. A writer/validator fusion is only a
follow-up hypothesis: parsing the pre-compaction `after` bytes does not by
itself prove that the emitted compacted bytes have the same semantics.

A retained candidate would need a structural proof that covers compacted
whitespace, `xml:space`, entity/general-reference handling, namespace and MCE
behavior, malformed-input errors, style and metadata limits, formula/value
ordering, web binding validation, original-byte rules, and post-write output
publication. Full output validation and readback remain required. The 0471
early buffer-release experiment was already rejected for lack of a practical
memory result. No vector-release work is queued here without a new,
operation-local reason; it is not a substitute for the larger changed-output
attribution.

The 0512 simulated instruction ceiling supports this ordering: source-store
parsing, worksheet rewrite, changed-output validation parsing, and compaction
are all large direct children, while shared-formula resolution and plain-cell
tag work are small. A tiny address/tag shortcut should not become the next
scope merely because it is easier to qualify.

## Option B: cold semantic no-op or memory admission fails

The current transaction selects `store_with_layout` from requested cell, row,
or column maps before semantic effectiveness is known. A cold same-value edit
therefore parses the semantic store, scans and allocates the full layout, then
discovers that projection is ineffective and drops the layout. The parser's
raw cells and the layout are live together. The hot same-value path is already
protected by the cache-hit branch; it does not answer the cold case.

There is one credible structural route to preserve fusion work elimination:
make fusion conditional on an **effectiveness proof that exists before the
fused pass starts**, tied to the same immutable source identity. Examples
would be a private transaction input carrying a source-checked effective
marker, or an operation whose source state was already established by an
earlier retained store query. Such a marker must be invalidated by any source
version or allocation-identity change.

The current public edit path does not carry that proof for a cold ordinary
cell/row/column action. A gate that first scans the target to discover whether
the action changes state and then restarts the fused reader would redo source
parsing. A gate that retains all pre-activation events to reconstruct the
layout would reproduce the memory overlap under another representation. A
gate that starts the scanner only after the target is found cannot rewrite the
prefix without a second scan or a complete event log. These forms do not meet
the stated requirement of avoiding speculative cold-no-op cost without
redoing parse or abandoning the work elimination.

Accordingly, if the measured cold no-op or incremental peak fails, use this
decision rule:

* retain a narrow, source-checked fusion gate only if a real pre-pass
  effectiveness proof is available and its changed-edit benefit remains
  measurable;
* otherwise keep the separate semantic parse and snapshot scan for this
  transaction path, with the current hot-store no-op behavior unchanged; and
* do not switch to a small cell-address/tag micro-optimization just to obtain
  an easy passing result. 0512 explicitly ranks the larger parser/Layout work
  above that fallback.

The safe fallback retains semantic-before-snapshot error order, x14ac/MCE
precedence, action-projection precedence, source-token checks, exact no-op
bytes/patches, and all existing rewrite validation. A failed performance gate
does not authorize relaxing any of those contracts.

## Option C: candidate performance is rejected after correctness passes

If the focused suite is green but the cold no-op or memory gate rejects the
current breadth, the source-fusion branch should stop at the measured
prototype. The separate changed-output path remains a viable XLSX target
because it runs only after an effective output exists and therefore has no
cold semantic-no-op admission problem. It still needs its own operation-local
profile and output-semantic proof; it must not be treated as an automatic
replacement for source fusion.

## Next substantiated OLE2/OOXML hotspot if source fusion is rejected

The strongest next XLSX scope is the **changed-output validation and
compaction path**. 0512 attributes 25.77% of direct commit instructions to
the changed worksheet validation parse and 20.38% to changed XML compaction.
The follow-up should profile those two boundaries separately before choosing a
writer/validator shared-event candidate. This scope occurs after an effective
rewrite, so it directly avoids the cold same-value problem. The emitted-byte
semantic proof, style/error order, web validation, exact output preservation,
and post-write readback remain mandatory.

The strongest cross-format alternative is **DOCX publication CPU
attribution**. 0496 finds publication as the largest named normal phase in
all 16 children, but its phase clocks are wall time and do not identify CPU
work. 0493's delayed-provider benefit is concrete but local-source timing is
near neutral and accepted overfetch is 37.29%. Profile publication, source
service, and allocator/hardware counters separately before changing either
publication passes or read-ahead policy. No local-source default change is
supported by the current evidence.

0511 closes the earlier CFB capture question. Its operation-local profile and
matched sector-batched FAT extension are already complete: source-open
instruction references fall from 15,127,968 to 13,973,601, and `load_fat`
falls from 1,407,020 to 250,820 (−82.17%). The larger residual shares were
`SectorChainScratch::collect_exact` 37.03%, `OleFile::claim_sector` 16.46%,
`validate_physical_sector_layout` 13.17%, and `validate_stream_allocations`
13.08%. These walks enforce chain, ownership, overlap, physical-layout, and
allocation invariants. They do not supply a safe repeatable elimination
candidate. The small repeated directory walk was only about 1.45% in the
retained attribution and remains deferred. Do not queue a repeated CFB profile
or remove those checks without a new concrete invariant-preserving
representation and reason.

## Guardrails carried into either branch

The 0499 ordered-Part scheduler flags remain active evidence: many-small
owned width-4 reused batch p50 is 186.04 microseconds versus 95.86 serial,
warm-file width-4 is 248.77 versus 182.04 serial, and delayed width-8 p99
rises 30.79% despite a 3.67% median reduction. Any later scheduler work still
needs operation-local wave/fence/queue/join attribution and must keep
operation-local workers, explicit stack/channel admission, ordinal error
selection, cancellation/source fences, and monotonic `Work`/`InputBytes`.
No hidden pool or admission-policy change is implied by this XLSX decision.

The 0500 managed paragraph-batch flags also remain active: owned p128 K=1
lifecycle p50 is +7.34%, p512 K=8 warm-file RSS is +5.77%, and the associated
publication p50 observation is +15.23%. They remain retained in any later
DOCX publication comparison and are not explained away by this XLSX result.

## References and stopping rule

The decision is grounded in [0511 CFB FAT reservation](../../changes/0511-cfb-fat-entry-reservation.md),
[0511 CFB source review](../change-0511/cfb-source-review.md),
[0511 profile scope review](../change-0511/profile-scope-review.md),
[0512 commit attribution](../../changes/0512-xlsx-commit-attribution.md),
[0513 operation allocation](../../changes/0513-xlsx-operation-allocation.md),
[0513 follow-up](../change-0513/follow-up.md), [0514 design review](design-review.md),
[0514 implementation review](implementation-review.md), and the [0510 OLE2/OOXML
priority queue](../change-0510/ole2-ooxml-priority-review.md).

Stop the fusion branch at the first failed correctness, cold no-op, memory,
or matched-benefit gate. If the branch passes, profile changed-output
validation/compaction next and choose a candidate only from its operation-local
evidence. If it fails without a valid structural gate, choose between that
changed-output XLSX scope and DOCX publication attribution; the CFB 0511 work
is already complete, and the small fallback work remains queued.
