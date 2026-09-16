Four log paragraphs for change 0660, for the coordinator to merge into the
shared program logs. Each block below is written in the style of the newest
section of the file it names.

## For `HOTSPOTS.md`

## 0660 — the DOCX edit's last whole-document pass is gone, and the thing that now bounds it is the publication audit

Record: [0660](0660-docx-compaction-policy.md).

**Change [0591](0591-docx-edit-single-scan.md)'s part (c) is implemented, and it
takes the second of the two remaining whole-part scans with it.** 0591 left an
ordinary one-paragraph edit and save scanning the main document twice: once for
the snapshot the caller asked for, once for `compact_changed_document_xml` and
the `with_rewritten_xml` rescan that followed it. Change 0652's decision 10
authorized the contract that blocked the second one, and it is now a public
policy: `litchi_docx::document::CompactionPolicy`, carried by `Edit`, defaulting
to `PreserveUnmodified`, with `WholeDocument` as the opt-in that reproduces the
base leg byte for byte. Under the default, `compact_changed_paragraphs` walks
the source and projected paragraph layouts in lockstep, compacts only the
paragraphs whose bytes differ — each alone, as a single closed root — and
splices them back through 0591's derived-layout route, so the whole-document
rescan disappears with the whole-document compaction. Per lifecycle of the three
harness shapes: `compact_changed_document_xml` **62,549,564 Ir → below the isolation pair's
own resolution**,
`scan_document_with_context` **6 calls → 3**, `Edit::commit` **133,165,034 Ir →
45,397,018**, whole iteration **572,582,252 → 484,151,439 (−15.44%)** on
`docx_semantic_one_edit_save` and −14.99% on the one-percent case; paired p50
**−34.14% / −30.00% / −16.61%** (large / medium / tiny) against an A/A floor of
at most 1.69% in the same window, repeating to within 2.3 points across the
three windows run. Isolated to the commit region alone, a
probe puts it at **−56.83% / −66.28% / −67.49%** at 24, 200 and 10,000
paragraphs. 0591's last named limitation goes in the same change:
`Package::document_snapshot` now hands `Snapshot::from_shared_xml` the part's
`Arc` instead of copying the blob, so `Snapshot::from_xml` falls from 3 calls to
0 per exact-no-op lifecycle; the no-op scenarios move −1.07% to −2.12% at p50
against an A/A floor reaching 1.69%, so the copy is **not resolved by timing**
and no latency figure is claimed for it. **The new hotspot is the
gate, not the compaction.** Preservation republishes the producer's own bytes,
and the OPC writer's authored-XML compactness contract refuses those bytes for
**54 of the 55 openable DOCX fixtures in this repository**, 53 of them on a line
break between the XML declaration and the root element. The default therefore
consults `xml_minifier::audit::verify_authored` — the writer's own auditor,
under the writer's limits — and falls back to the unchanged whole-document
route when it says no. A short-circuit on the one refusal that holds
unconditionally, character data outside the root element, decides those 53
fixtures from about sixty bytes of prologue, so the fallback costs **1,695
instructions on a 24-paragraph document and nothing measurable on a
10,000-paragraph one** instead of the 44.9 M Ir a full audit would have cost.
Change 0652's decision 2 (queue row 2) is what moves that contract; when it
lands, `publication_accepts_preserved_xml` is the single function that has to
change, and this record's census is the before-picture. OLE2/OOXML remain
active; ODF is deferred until completion and iWork excluded.
[Record and limitations](0660-docx-compaction-policy.md);
[retained evidence](results/change-0660/README.md).

## For `GOAL_AUDIT.md`

## 0660 — preservation by default reaches inside the part it edits, and stops exactly where the writer's contract does

Record: [0660](0660-docx-compaction-policy.md).

`docs/GOAL.md` puts lossless preservation and correctness above speed, and ADR
0006 makes preservation the default rather than an option. Until this change
that default stopped at the part boundary: an ordinary DOCX edit preserved every
part it did not name, and then re-serialized every paragraph of the one part it
did — stripping whitespace, re-escaping attribute values and normalizing
attribute spelling in paragraphs the caller never touched. 0591 named that as an
owner question rather than a defect and froze it; 0652's decision 10 answered
it. The audit distinction worth recording is that **the change is bounded by a
contract it deliberately did not move**. Preserved bytes must still satisfy the
package writer's authored-XML compactness audit, which is change 0652's decision
2 and another record's work in this wave; rather than widen its own default past
that, 0660 asks the writer's own auditor first and takes the previous route when
the answer is no. The consequence is stated plainly and is the record's main
weakness: **the preserving route is reachable on one of the 55 openable
fixtures today**, so the contract this change exists to deliver is exercised by
that fixture and by a constructed witness, while the measured saving comes from
removing work on documents that already satisfy the writer. Two admissions, both
stated rather than buried. First, **one refusal class narrows**: a refusal that
existed only because whole-document compaction re-serialized *untouched* markup
— an `xml:space` value outside `default`/`preserve` in a paragraph the edit
never named — is not raised under the default policy. That is inseparable from
"leaves every unmodified paragraph untouched", the opt-in still raises it with
the identical error, and every refusal belonging to the snapshot itself — the
byte, node and depth limits, DTDs, processing instructions, unbalanced nesting —
is untouched because the snapshot scan is untouched. Second, **a source the gate
refuses now pays a little more than it did**: the short-circuit, +0.59% of the
commit region on the smallest shape and below measurement on the largest,
reported rather than averaged away. The oracle is a corpus census of 63
fixtures × ten aspects on both legs — the ordinary save, the exact no-op, the
managed insert-and-publish route, a one-paragraph rewrite, the source-versus-
published byte span and a reopen — and it is **byte-identical on every aspect
and every refusal string**, except the opt-in aspect the base leg cannot run.
No latency claim is registered. OLE2/OOXML remain active; ODF is deferred until
completion and iWork excluded.

## For `REPORT.md`

## 0660 — editing one paragraph of a Word document no longer rewrites all the others

Record: [0660](0660-docx-compaction-policy.md).

`litchi-docx`: when Litchi changed a single paragraph of a Word document and
saved it, it rewrote the document's entire internal XML from scratch —
squeezing out spacing and re-spelling quotation marks in paragraphs nobody had
touched. The file still said the same thing, but the bytes of untouched
paragraphs moved, and every save paid for a full rewrite plus a full re-read of
the document. The owner's decision for this wave was to make that a setting and
to default it off. It now is: a new public `CompactionPolicy` on the edit, whose
default compacts **only the paragraphs the edit actually changed** and
republishes every other byte exactly as the original program wrote it, and whose
opt-in (`WholeDocument`) reproduces the old behaviour byte for byte for anyone
who wants it. Removing that rewrite and the re-read behind it makes a
one-paragraph edit-and-save **34.1% faster on a 10,000-paragraph document, 30.0%
on a 200-paragraph one and 16.6% on a small one** (median of 100 samples per
side, against a same-window noise floor under 1.7%, and repeating to within 2.3
points across three separate measurement windows); measured on the edit step
alone the saving is **about two thirds**. A second, smaller change lands with
it: taking a document snapshot no longer copies the whole main part — worth one
whole-document copy and one allocation per snapshot, which is too small to
separate from measurement noise on this machine and is claimed only as removed
work, not as a speed-up. **Nothing that gets written
changed.** Over all 63 Word-family test documents in the repository, checked ten
different ways each — plain save, empty edit, paragraph insert, paragraph
rewrite, and reopening the result — every single saved file is byte-for-byte
identical to what the previous version produced, and every refusal message is
word-for-word the same. The honest limitation is where the new default can
actually be used today: Litchi's package writer insists that any XML it
regenerates be written in a strict "compact" form, and **54 of the 55 openable
test documents fail that check** — almost all of them because their producer put
a line break after the XML declaration. Until the companion change in this wave
relaxes that check, Litchi quietly falls back to the old whole-document rewrite
for those files, at a cost of about 1,700 machine instructions to notice.
Nine new tests, including one that replays both settings over every fixture in
the repository; `cargo fmt`, clippy, 1,480 `litchi-docx` tests, rustdoc, the
feature-bearing facade suite, the 540-test performance harness and the
non-iWork gate all pass. `performance_claim: none`; the medians above are
reported as evidence beside their noise floor, not registered as a claim.
[Change and limitations](0660-docx-compaction-policy.md); [retained
evidence](results/change-0660/README.md).

## For `ADR_COMPLIANCE.md`

## 0660 — compliant; ADR 0006 chooses the default, and a contract this record consults rather than moves

Record: [0660](0660-docx-compaction-policy.md).

Change 0660 is compliant and amends no ADR — change 0652 assigns ADR amendments
to decisions 2, 4 and 5, not to decision 10 — but it is decided by two.
**ADR 0006 (preserve by default)** chooses which way round the new public
`CompactionPolicy` points: preservation is the default and whole-document
compaction the opt-in, because ADR 0006 makes preservation the thing a caller
gets without asking. It also settles how far preservation reaches: into the
edited part's untouched paragraphs, not merely to untouched parts. **ADR 0003
(edits are source-checked and reversible; a refusal is typed, never a partial
result)** is preserved: `commit()` returns the same `Commit`, the snapshot stays
immutable and cheaply shared, `Patch::apply`'s `same_source` still gates every
application, and the semantic readback after each rewrite is untouched. **ADR
0005 (mandatory validation; no leakage)** is preserved in both halves: no check
is removed, weakened, reordered or made conditional; `scan_document` runs on
exactly the inputs it ran on with `MAX_DOCUMENT_XML_BYTES`,
`MAX_DOCUMENT_NODES` and `MAX_DOCUMENT_DEPTH` at their values, identities and
enforcement points; every splice still has to satisfy 0591's
`preserves_body_child_shape`, and anything it declines falls back to the full
rescan, so no new path exists between unproven bytes and a snapshot. No archive
type, lock or executor reaches the public API; no new `unsafe` (the crate
remains `#![forbid(unsafe_code)]`); no new dependency (`xml-minifier` was
already a `litchi-docx` dependency); no ambient I/O and no global pool. **Two
rows are stated rather than left implicit.** First, the **breaking change**
0652's trade-off 1 authorizes: `Edit::commit`'s default output bytes change for
a source that satisfies the publication contract and is not invariant under
compaction, and with it **one refusal class narrows** — a refusal raised only
because whole-document compaction re-serialized untouched markup is not raised
for markup the default does not re-serialize. The opt-in raises it with the
identical error, no snapshot-level refusal moved, and the 63-fixture census
shows no refusal string differing on any route. Second, the **contract this
record consults but does not move**: the OPC writer's authored-XML compactness
audit refuses preserved bytes for 54 of 55 openable fixtures, and rather than
widen the default past it, `CompactionPolicy::PreserveUnmodified` asks that
auditor — the writer's own, under the writer's limits, so the two cannot
disagree — and falls back to the base leg's route. That contract is 0652's
decision 2, row 2 of 0651's queue, and belongs to another record in this wave.
Change [0629](0629-facade-docx-budget-test-bisect.md)'s managed budget contract
is unaffected: both splice routes produce `admission: None` and an owned storage
exactly as `with_rewritten_xml` did, so a compacting commit detached its managed
admission before this record as it does after it, and no `reserve_managed` call
site moved. Change [0650](0650-docx-editor-byte-order-mark-admission.md)'s
byte-order-mark carry is untouched — `compact_changed_document_xml` is not
modified — and none of 0650's four frozen follow-ups is resolved here.
`performance_claim: none`. OLE2/OOXML remain active; ODF is deferred until
completion and iWork excluded.
