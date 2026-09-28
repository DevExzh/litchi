# 0807 — current PPTX capture source review

This is a source-only review at `3624e73236` (`perf(xml): reject combined
candidate at public workflow benefit gate`). The 0806 XML iterator candidate was
rejected and the six production files were restored; this review proposes no
production edit, build, capture, timing result, allocation result, or savings
claim. The 0807 packet is the fresh CPU-attribution follow-up. Its 0791 lanes
are useful only for choosing what to count: the 0791 profile predates the
retained 0792 empty-attribute-tail shortcut and the 0800 duplicate-error parity
correction, and the 0798 census was diagnostic rather than a cost model.

The review read the accepted constraints in ADR 0001 (correctness and
preservation), ADR 0003 (immutable snapshots and atomic validation), ADR 0005
(profiles and statistical evidence before optimization), ADR 0006 (preserve
mode and typed refusal), ADR 0008 (verification gates), ADR 0013 (notes graph
ownership), and ADR 0032 (small payload-derived classification memos). The
relevant source history is the retained capture proof in `829bed6965`, the
memo and ownership fixes in `c4adb982f7` and `105a38bc7c`, and the rejected
0806 candidate at the current HEAD.

## Current reachable path

The opened-presentation path is ordered as follows.

```text
capture_internal
  PresentationPart::from_package                 main root/type preflight
  PresentationPart::slide_references             validated slide catalog
  Presentation::catalog                          per-view catalog memo hit
  capture_slides_inner                           one slide at a time
    validate_slide_relationship
    processed_xml_with_capture                   MCE projection
    root_name_from_xml                           first-root/name gate
    c_sld_name_from_xml                          producer-visible name
    root_conformance_from_processed              optional notes-root proof
  identity/name loop                             relationship and deferred error order
  load_index_with_slide_root_proofs               notes graph
    root_conformance(presentation)               classify presentation root
    scan_xml(presentation, conformance)           second presentation scan
    resolve_slide_root(proof or raw scan)         classify every slide root
```

The first part of this route is in
[`opened/model.rs`](../../../../crates/litchi-pptx/src/opened/model.rs:730),
the per-slide projection is in
[`presentation/package.rs`](../../../../crates/litchi-pptx/src/presentation/package.rs:65),
and the proof is made in
[`parts/slide.rs`](../../../../crates/litchi-pptx/src/parts/slide.rs:1131).
`capture_internal` does not enter notes loading until every mandatory slide
root, identity, and name check has passed; a deferred first name error returns
at
[`opened/model.rs:796`](../../../../crates/litchi-pptx/src/opened/model.rs:796)
before the notes call at line 818. A malformed root returns earlier from the
per-slide projection. This is why the relative position of root, name, and
notes failures is observable even though the final successful snapshot has no
proof objects in its public API.

The notes index is in
[`notes/package.rs`](../../../../crates/litchi-pptx/src/notes/package.rs:426).
It scans the presentation to collect both `r:notesMasterId` and local-name
`sldId` inventories, then validates each relationship and conformance. A proof
is used only when the expected catalog length equals the scanned notes
inventory and its raw witness has the same pointer and length as the current
slide payload (`resolve_slide_root`, lines 703–712). A mismatched witness falls
back to the ordinary raw scan. An exact proof with `None` conformance preserves
the generic `invalid sld root or namespace` refusal.

The following conditions are actually reachable and must be included in any
candidate model:

* `proofs.try_reserve_exact(references.len())` can fail at
  [`presentation/package.rs:79`](../../../../crates/litchi-pptx/src/presentation/package.rs:79).
  Capture then continues without a proof vector and notes validation rescans
  slide roots from raw bytes.
* A successful proof vector is disabled for later slides after the first
  `None` conformance at lines 105–111. The current invalid proof remains in the
  prefix, while later positions use raw fallback. This is a safety fallback,
  not a reason to trust a shorter vector as a complete inventory.
* A fresh opened capture has no parent `SlideRootMemo`, so every valid named
  slide takes the complete processed-root classification path when collection
  remains enabled. A commit can hit the parent memo only for the exact source
  allocation and length; a new allocation with equal bytes is a miss. The memo
  retains successful classifications only, never refusals, and debug/test
  builds rederive hits.
* `root_conformance` tries Transitional and then Strict and masks both failed
  scans to the generic root refusal. `root_conformance_from_processed` has the
  same retry with the separate raw-length check. The 16 MiB notes-root ceiling
  and the 64 MiB PresentationML part preflight are distinct reachable gates.
* A foreign-namespace `sldId` lookalike makes the notes inventory differ from
  the ordinary presentation catalog. The proof hint is then disabled and the
  missing relationship error remains the notes-graph result. It must not be
  “fixed” by trusting the catalog’s narrower namespace interpretation.

## Why deleting slide validation is unsafe, while moving the optional proof is a valid hypothesis

There are two distinct checks in the slide projection. The mandatory
`root_name_from_xml` check at
[`parts/slide.rs:1137`](../../../../crates/litchi-pptx/src/parts/slide.rs:1137)
returns its root error directly. The complete scan at
[`parts/slide.rs:1141`](../../../../crates/litchi-pptx/src/parts/slide.rs:1141)
produces an optional `SlideRootProof`:
`root_conformance_from_processed` returns `Option<Conformance>`, and its failed
classification is stored as `None` rather than propagated from
`finish_from_processed`. ADR 0032 requires the notes graph to classify every
slide root, so the proof is useful because it performs that classification
while the capture’s processed XML is still available. Removing only this
optional proof scan and leaving the notes loader unchanged moves the same
classification to
[`notes/package.rs:469`](../../../../crates/litchi-pptx/src/notes/package.rs:469);
it does not remove a pass. It can also lose the capture-to-notes MCE reuse and
repeat processing, which is a work hypothesis for profiling. Removing the
mandatory root preflight as well, or using proof absence to skip the later
notes classification, would violate the notes owner contract. The optional
move alone is not proof of a refusal-order change: the existing raw fallback
is designed to preserve the notes-stage generic refusal.

The mandatory root/name loop still establishes the observable ordering before
that optional distinction matters. Existing oracles freeze this sequence:

* [`opened/tests.rs:329`](../../../../crates/litchi-pptx/src/opened/tests.rs:329)
  requires a later bad slide root to beat an earlier duplicate catalog/name
  condition, while a duplicate main-part catalog is rejected before slide
  projection.
* [`opened/tests.rs:357`](../../../../crates/litchi-pptx/src/opened/tests.rs:357)
  and [`:473`](../../../../crates/litchi-pptx/src/opened/tests.rs:473) require
  an earlier deferred name error to beat malformed notes XML, while a later
  slide-root error beats both the name replay and notes graph.
* [`opened/tests.rs:390`](../../../../crates/litchi-pptx/src/opened/tests.rs:390)
  keeps mixed slide conformance as a notes-stage refusal, after the slide
  projection accepts either root dialect.
* [`opened/tests.rs:503`](../../../../crates/litchi-pptx/src/opened/tests.rs:503)
  requires a foreign inventory mismatch to fall back to the current raw scan
  and retain the missing-relationship error and name-error deferral.
* [`opened/tests.rs:526`](../../../../crates/litchi-pptx/src/opened/tests.rs:526)
  keeps the generic 16 MiB notes-root refusal distinct from the earlier 64 MiB
  part preflight.

The retained 0693 evidence gives a related historical warning: a prior
capture-local proof design changed the cost of late refusals because earlier
slides were fully validated first. That result is not a current 0807 timing
claim and does not show that disabling this optional hint changes refusal
ordering. It does show why any fusion must separately test mandatory root/name
errors, optional proof `None`, and the raw notes fallback. The credible work
removal is either to eliminate a genuinely duplicate scan, or to fuse the
optional classification with another required pass while returning the same
conformance, inventory, and fallback behavior.

## Candidates and ranking for the fresh profile

### 1. Measure a slide scanner that also projects `cSld@name`

This is the highest-upside source hypothesis to measure because the retained
0791 attribution places the per-slide `scan_processed_xml` path above the
presentation-index path. It is a hypothesis about redundant event decoding,
not a projected speedup. Today `finish_from_processed` first calls
`root_name_from_xml`, then `c_sld_name_from_xml` (which delegates to
`namespace::presentation_name`), and only then performs the complete
`root_conformance_from_processed` scan. The full scan already visits every
element and checked attribute in the processed slide. A private combined
projection could return the root classification and the producer name from
that one event walk, dropping the separate `cSld` prefix walk while retaining
the current proof result.

This candidate is not ready for an edit. `presentation_name` accepts either
PresentationML dialect, recognizes the conventional unresolved `p` prefix, and
uses `unqualified_attribute_value`, whose duplicate and malformed-attribute
errors are observable. The current proof is deliberately gated by
`name.is_ok()`. A fused helper therefore has to preserve all of these cases:

* root-name rejection remains first, including the existing handling of
  declarations, comments, whitespace, non-element events, and invalid root
  namespaces;
* a duplicate or malformed `cSld@name` remains the deferred first name error
  and must not be replaced by the proof scanner’s generic root error;
* a missing `cSld@name` still falls back to the part name, and an explicitly
  empty value remains an empty name;
* Transitional-then-Strict classification, raw and processed byte limits,
  node/depth/attribute budgets, DTD/PI/CDATA refusal, relationship inventory,
  and generic failure masking remain unchanged; and
* a name error must prevent that slide’s proof from being treated as valid,
  while later-slide root errors retain their current position in the outer
  capture loop.

The profile must first report `cSld` byte offsets, name-parser calls and bytes,
proof-scanner calls and bytes, and conformance retry outcomes. If the name
parser normally stops at a tiny prefix, a fused implementation may have too
little ceiling to justify its semantic risk. If the scanner dominates and the
prefix is material, the differential oracle should compare every returned name,
conformance, refusal string, and proof fallback against the current route.

### 2. Return the successful presentation `XmlScan` with its conformance

There is a mechanically redundant pair at
[`notes/package.rs:435`](../../../../crates/litchi-pptx/src/notes/package.rs:435):
`root_conformance` scans the complete presentation XML to find the dialect,
then `scan_xml` scans the same payload again with that dialect to obtain the
`XmlScan` inventory. A private helper can try Transitional and Strict in the
same order and return `(Conformance, XmlScan)` from the first successful scan.
`root_conformance` can remain a conformance-only wrapper for other callers.

This is the safest concrete redundant-work seam, but not necessarily the
largest one. It must preserve the generic error mask when both dialects fail,
the raw/processed byte limits, all inventory fields, and the ordinary notes
callers. The 0791 frame-pointer rows put
`load_index_with_slide_root_proofs` at only 8/5 nested owner samples, so the
current evidence ranks this below the per-slide scanner; those counts predate
the current source and are not a savings claim. Fresh 0807 counts should
measure presentation bytes and both conformance attempts before any edit.

The focused oracle is already present in
[`notes/codec.rs:523`](../../../../crates/litchi-pptx/src/notes/codec.rs:523)
for dialect retry, generic masking, processed classification, and limits. The
notes graph tests must additionally compare relationship inventories and
strict/transitional graph results through `load`, opened capture, and the
`put_changed` validation path at `notes/package.rs:1225`.

### 3. Combine the opened catalog pass with the notes inventory pass

`PresentationPart::slide_references` at
[`parts/presentation.rs:114`](../../../../crates/litchi-pptx/src/parts/presentation.rs:114)
fully parses the main part for the validated ordered catalog. The notes loader
then scans the same presentation payload again to collect its own
`notesMasterId` and local-name `sldId` inventory. The `Presentation::catalog`
`OnceLock` avoids a second catalog parse within the view, but it cannot answer
the notes loader’s separate graph contract.

A capture-only combined scanner could return the validated catalog and notes
inventory together, then pass the retained result into the notes loader. This
has a larger semantic surface than candidate 2: it must retain the catalog’s
direct-list shape, uniqueness, relationship-target and content-type checks,
while still collecting foreign local-name `sldId` entries that the ordinary
catalog intentionally ignores. It must also keep catalog errors before slide
projection and notes errors after identities and names. The foreign-inventory
tests at `opened/tests.rs:503–522` are a mandatory guard. Ordinary notes loads
and `put_changed` cannot inherit a capture-only result without an explicit
ownership and lifetime design. Measure catalog and notes-scan bytes first; no
benefit follows from their both being XML passes.

### Excluded from this review

The 0806 checked-attribute iterator candidate is rejected at the public
workflow gate and is not retried under another name. The 0792 exact empty-tail
shortcut is already retained at `notes/codec.rs:474`. A parsed-XML cache or a
larger notes inventory cache would violate ADR 0005/0032’s retention boundary;
the existing MCE and slide-root memos retain only bounded values tied to exact
payload allocations. Unchecked attribute parsing, skipping root validation, or
trusting the ordinary catalog in place of the notes inventory is not a
candidate.

## Future fusion prerequisites (outside the frozen 0807 packet)

The frozen 0807 protocol supplies owner-scoped native, Callgrind, and
frame-pointer CPU attribution. It does **not** collect per-function byte
lengths, `cSld` offsets, conformance retry outcomes, notes-present refusal
fixtures, or recapture memo/proof counters. None of those facts is claimed by
this packet or by this source review. A later fusion experiment would need to
collect the following before implementation:

* call count and input bytes for `root_name_from_xml`,
  `presentation_name`, `root_conformance_from_processed`,
  `scan_processed_xml`, `root_conformance`, `scan_xml`, and
  `PresentationPart::slide_references`;
* Transitional success, Transitional failure followed by Strict success, and
  both-dialect failure for presentation and slide roots;
* fresh capture versus commit recapture, including exact slide-root memo hits,
  misses, invalid classifications, and proof-vector reservation fallback;
* notes-present versus notes-absent packages, inventory match versus foreign
  inventory mismatch, and name/root/notes refusal fixtures; and
* per-slide processed/raw lengths and the position of `cSld`, so a prefix
  projection is not mistaken for a second full scan.

Callgrind can bound local work, native controls qualify operation latency, and
the frame-pointer lane can qualify owner ancestry. Their nested counts must not
be added as independent shares. The source review makes no conclusion until
the current 0807 outputs answer whether either redundant pass is material.

## Existing semantic oracles

Any future candidate must keep the current tests and add only the differential
cases needed for the fused route. The relevant existing coverage is:

* `opened/tests.rs:329–417` for catalog/root/name/notes precedence, strict MCE,
  and mixed conformance;
* `opened/tests.rs:421–522` for exact proof identity, reordered equal-length
  fallback, invalid-proof refusal, and foreign inventory fallback;
* `opened/tests.rs:526–567` for the separate 16 MiB and 64 MiB limits;
* `opened/slide_root_memo_tests.rs:182–320` and `:653–676` for one proof per
  scanned slide, commit reuse, equal-byte allocation misses, refusal parity,
  and notes-owning commits;
* `notes/codec.rs:523–706` and its buffered differential oracle for dialect,
  parser, malformed-attribute, and limit behavior; and
* `notes/tests.rs:320–490` for complete graph, relationship, orphan, and
  atomic-mutation validation.

These tests cover successful snapshots and typed refusals; byte preservation
and proof ownership remain part of the oracle. A source-level simplification
must not change which malformed or unknown content is preserved, rejected, or
reported first.

## Disposition

No production candidate is selected from this source review. The first
measurement question is whether the per-slide scanner can safely subsume the
name projection; the first mechanically bounded implementation candidate is
the presentation `(Conformance, XmlScan)` helper. The catalog/inventory fusion
is a larger follow-up only if current attribution shows material duplicate
work. The whole slide proof scan remains required unless one of those fused
passes returns its exact classification and preserves every refusal and
fallback rule above.
