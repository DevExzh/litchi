# Log sections for change 0607

Four paragraphs for the coordinator to merge, one each into `HOTSPOTS.md`,
`GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, in the style of their
newest sections. This record touched none of those files.

---

## For `HOTSPOTS.md`

## 0607 — the PPTX slide regeneration is unreachable, and the authored save's whale is the deflate encoder

Change 0587's item SAVE-3 ranked `materialize_presentation`'s rebuild of every
slide under the scenario "PPTX opened-document update/save". Measured, that
scenario does not exist: `Package::mutable_pres` is assigned `Some(..)` in
exactly one place, `Package::new()`, and every byte ingress funnels through
`from_opc_package_with_provenance`, which assigns `None`; all 78 `.pptx`
fixtures under `test-data` and both litchi-authored decks are refused by
`presentation_mut` with the same typed `UnsafeEdit`. The route that does serve
opened decks, `opened::patch::apply`, already applies a per-part delta and
renumbers nothing, and an unedited opened save is already exact-source
passthrough, byte-identical on 78 of 78. The one reachable route is an authored
deck materialized more than once, where SAVE-3's cost model is wrong by
construction: with no source archive the publication plan has no `Copy` action,
so all 161 members of a 50-slide deck are serialized into the manifest, audited
and deflated on every save whatever the model does. The whole materialization is
3,665,367 Ir = 2.28% of a 161,032,882 Ir edit-and-save, and 10.00% (50 slides)
to 11.18% (200 slides) of its wall-clock p50 against +0.43% and +0.06% p50 A/A
floors. The same profile re-points the real cost: `DeflateEncoder::new` runs 161
times per authored save at 444,764 Ir each, 44.5% of the save's instructions —
survey item §2.4 / SAVE-6 in `crates/soapberry-zip/src/office.rs`, which change
0476's `OwnedDeflateState` reuse does not cover. SAVE-3 should be re-ranked as
an authored-path item, below SAVE-6. No production change and no speedup is
claimed. [Design, gates and falsification](0607-pptx-authored-slide-regeneration-design.md);
[evidence](results/change-0607/README.md).

---

## For `GOAL_AUDIT.md`

## 0607 — PPTX opened-document CRUD has one route, and the example that takes the other is broken

The audit's standing P1 row "finish source-backed CRUD adoption across formats"
gains a measured boundary on PPTX. There are two writer surfaces in
`litchi-pptx`: the mutable `MutablePresentation`, reachable only from
`Package::new()`, and the `opened::Transaction` capture-and-patch route,
reachable only from an opened package. They are mutually exclusive by
construction, and the census confirms it on the whole corpus: 78 of 78 `.pptx`
fixtures open and 78 of 78 refuse `presentation_mut`, while an unedited save of
each is byte-identical to its source. Opened-document PPTX CRUD is therefore
entirely `opened/`'s (changes 0590 and 0598), and the mutable writer is an
authoring surface only. One in-tree consumer has not noticed:
`crates/litchi/examples/office_crud_demo.rs:289-290` performs its PPTX UPDATE
step as `Package::open` followed by `presentation_mut`, which cannot succeed, so
the CRUD demonstration fails at that step. Reported, not fixed: the fix is to
port the step to `opened_presentation_transaction`, which belongs to the
`opened/` owner. Nothing in the repository materializes an authored presentation
more than once, so the authored-save opportunity this record freezes has no
caller and no harness selector to be measured against, and that is its admission
gate. `performance_claim: none`; `claim_authorized: false`. OLE2/OOXML remain
active; ODF is deferred until completion and iWork excluded.

---

## For `REPORT.md`

## 0607 — the authored PPTX save, measured end to end for the first time

No record had profiled an eager PPTX save. On an authored 50-slide deck (161
members, 91,537 bytes) a callgrind isolation pair at N=1 and N=6 puts one
edit-and-save at 161,032,882 Ir, of which `PackageWriter::to_bytes` is 97.8%,
`PhysPkgWriter::write` 89.7%, a fresh `DeflateEncoder` per member 44.5% (161
constructions at 444,764 Ir), the deflate itself 27.2%,
`PublicationPlan::from_package` 7.9%, `verify_authored` 6.9% (161 calls) and the
presentation materialization 2.20%. Four paired legs in A1 B1 B2 A2 order, 30
samples each after three warmups, pinned to CPU 27, put the materialization at
326.2 µs of a 3,261.0 µs p50 at 50 slides and 1,098.4 µs of a 9,827.6 µs p50 at
200, against A/A floors of +0.43% and +0.06% p50 — five times its instruction
share, because callgrind counts the deflate encoder's state zeroing per byte.
Two byte oracles bound any future change: re-materializing an unchanged model
reproduces the previous archive byte for byte, at 3 and at 50 slides, and a
one-slide edit already leaves 160 of 161 members byte-identical, changing only
`ppt/slides/slide1.xml`. The relationship renumbering the survey worried about
is already a no-op, because `Relationships::get_or_add` returns an existing id
for an unchanged type and target and `next_r_id` fills the lowest gap. No
production code changed and no real producer file was measured, because none can
reach this path. See [Change 0607](0607-pptx-authored-slide-regeneration-design.md);
`performance_claim: none`.

---

## For `ADR_COMPLIANCE.md`

## 0607 — preserving an unmodified slide is the more correct behaviour, under seven gates

The frozen design keeps unmodified slides' parts in place instead of deleting
and rebuilding them, and ADR 0006's preservation default argues for it: not
rewriting a member whose value has not changed is what preservation means, and
the measurement shows the rewrite already reproduces the member byte for byte,
so nothing a consumer observes moves. ADR 0005 is not relaxed. No limit moves
and no refusal moves: the fast path performs a strict subset of the fallible
calls today's path performs, and every one it skips is skipped only after a
pointer-identity proof — `Arc::ptr_eq` between the live part's `blob_arc()` and
the `Arc` the last materialization installed — that the value it would have
produced is already in the graph. This is change 0593's direction of proof: a
match proves "unchanged", a mismatch proves nothing and costs the existing path,
and `OpcPackage::clone` clones the `Arc` rather than the bytes, so the proof
survives the rollback snapshots `flush_presentation` and `edit_raw` take. Steady
state grows by one pointer pair per slide, not a second copy of the deck. Seven
gates, all evaluated before any mutation, keep the fast path from inheriting an
out-of-band graph edit: the presentation's own `modified` flag must be clear
(which also refuses every first save and closes the `duplicate_slide` clone
hazard), the slide and notes parts must be exactly the recorded set with the
recorded content types and `Arc`s, their relationship triples must be exactly
the recorded ones, and the presentation part's slide relationships must be
exactly the recorded id-and-target pairs. A document with gaps or non-canonical
slide names cannot arise on this path, because only this function ever creates a
slide part; if one is introduced out of band the gates refuse and today's
renaming path runs unchanged. `verify_authored` still runs over all 161 members
of every authored save, because an authored package has no provenance that could
exempt one. No ADR is amended and no ADR clarification is proposed.
[Change 0607](0607-pptx-authored-slide-regeneration-design.md);
`performance_claim: none`.
