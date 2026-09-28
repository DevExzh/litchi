# 0811 source/path review: repeated slide projection work

This is a read-only source review after the retained 0810 direct-event change.
It proposes one bounded experiment and no production implementation. No Cargo,
benchmark, workload, or profiler command was run for this review. The current
HEAD is `2894fbd628` (`perf(pptx): read notes XML events without NsReader
wrapper`); the current `notes/codec.rs` SHA-256 is
`466e504588843a0b2538fc25ce3378b6fb9bf13aa9b0d74efebcb73ca0ebab5e`. The
0810 and 0811 architecture-input manifests are byte-identical, and all 35
recorded goal, taxonomy, accepted-ADR, and architecture files still hash to
their 0810 values. The unrelated worktree changes were left untouched.

## Current path

The opened PPTX capture has several distinct observations of the same XML
bytes:

1. `capture_internal` constructs a `PresentationPart`, reads its slide
   catalog, and then captures each slide through
   [`PresentationPart::from_part`](../../../../crates/litchi-pptx/src/parts/presentation.rs)
   and [`PresentationPart::slide_references`](../../../../crates/litchi-pptx/src/parts/presentation.rs).
   Those are separate presentation projections. The MCE capture can retain a
   transformed slide projection, but it does not merge the presentation
   root and catalog readers.
2. For each slide,
   [`SlidePart::finish_from_processed`](../../../../crates/litchi-pptx/src/parts/slide.rs)
   receives one already-processed XML value. It first calls
   `root_name_from_xml`, then `c_sld_name_from_xml` (which delegates to
   `namespace::presentation_name`), and then, when the name projection has no
   error, calls `notes::root_conformance_from_processed` over the same
   processed bytes. The first two are namespace-reader projections; the last
   is the complete bounded notes scanner.
3. The notes graph later calls
   [`load_index_with_slide_root_proofs`](../../../../crates/litchi-pptx/src/notes/package.rs).
   It still scans the presentation once for conformance and once for its
   inventory, and validates the notes master and theme with one scan each.
   For a slide whose proof has the same raw allocation, `resolve_slide_root`
   consumes the proof rather than rescanning the slide. That proof/memo
   cross-check is a required safety boundary.

The generated public fixture always creates a notes master and notes theme,
while the writer creates a notes-slide part only when `slide.has_notes()` is
true. A no-notes fixture therefore still exercises the complete slide-root
proof and master/theme validation. The absence of a notes-slide relationship
is not evidence that a slide XML root may be trusted or skipped.

## Measured evidence from the retained 0810 profile

The sealed after Callgrind publications provide the following measured facts;
the values are guest-instruction attribution for the exact capture wrapper,
not native time or a universal workload claim:

| Path | Evidence in the large capture |
| --- | ---: |
| `SlidePart::finish_from_processed` → `scan_processed_xml` | 100 complete slide scans |
| `scan_processed_xml` event loop | 282,612 `Reader::read_event_impl` calls |
| `scan_processed_xml` self cost | 22,295,758 Ir |
| `scan_processed_xml` inclusive edge from slide capture | 317,358,177 and 317,403,524 Ir in the two after publications |
| `scan_xml` calls from the notes graph | 4 (presentation root/inventory plus master and theme) |

The direct 0810 reader change removed the scanner's
`NsReader::process_event` edge, but it did not remove the complete scans,
namespace scope transitions, or notes attribute validation. The existing
known-URI fast path in [`notes::resolved`](../../../../crates/litchi-pptx/src/notes/mod.rs)
is already present and is not a new opportunity. Likewise, the 0810
allocation reports have equal call and byte totals before and after; this
review does not claim an allocation reduction from any source observation.

The profile also records a large SHA-256 guest-Ir row in the broader capture
wrapper. That is a separate native-diagnosis question, especially because
guest `sha2::compress256` attribution is not a native cycle result. This memo
does not recommend changing the complete-package fingerprint or its
validation authority from that row alone.

## Inferred source opportunities

The bounded source repeat is the two short projection walks before the
complete proof. `root_name_from_xml` allocates a local name byte vector and a
`String` only to compare it with `"sld"`; `presentation_name` walks until the
first matching `cSld` or EOF and may own a producer name. The complete scanner
then parses the same processed bytes again. These readers generally return
early. In the retained after profile the namespace-reader edges attributed to
the slide path are only 400 calls from `root_name_from_xml` and 300 calls from
`SlidePart::finish_from_processed` (the inlined/name-projection path), versus
282,612 event reads in the 100 complete proof scans. The profile therefore
does not establish a large native gain from fusing the short projections, and
it does not prove that the root-name temporary dominates allocator work. Those
are inferences requiring a candidate measurement.

The presentation root/conformance scan followed by a second presentation
inventory scan is another real source repeat, but only four `scan_xml` calls
appear in the large profile. It is a lower-scale follow-up and should not be
bundled into the slide experiment. There is no source or profile evidence here
that `Part::blob()` calls trigger repeated decompression; the current path
borrows package-owned bytes, and archive decode caching belongs to the OPC
layer. A repeated-call count alone is not a decompression measurement.

## Safety boundaries and rejected shortcuts

The complete slide proof cannot be removed merely because a slide has no
notes relationship. The focused memo tests require one classification per
captured slide, reject CDATA and an unbound prefix in slides that otherwise
pass root/name projection, compare memo-assisted and cold refusals, and require
exact allocation-identity keys. `load_index_with_slide_root_proofs` must keep
its presentation conformance comparison, relationship checks, orphan checks,
resource limits, and master/theme validation.

Any reuse must also preserve the current order: root validity is observed
before the producer-name projection; a name/parser error prevents a notes
proof from being published; when proof collection is admitted, a successful
name projection is followed by the full bounded proof; and a proof hit is
accepted only for the exact raw allocation it classified. The scanner's
Transitional-then-Strict retry and
generic root-error masking, XML grammar/refusal behavior, depth/node/
attribute limits, MCE selection, unknown markup and namespace preservation,
and the no-op/memo/rebind contracts remain unchanged.

The rejected 0806 candidate is not a reason to retry a generic attribute
iterator or checked-attribute optimization. Its combined `xml_attributes`
change reduced large-capture allocation calls from 72,106 to 10,788 and
allocated bytes from 4,767,939 to 853,507, yet the frozen workflow gate still
rejected it: large capture was 1.853% slower, large lifecycle was 3.105%
slower, and no native capture or lifecycle row met the required benefit. The
current 0810 allocation comparisons are equal as well. The experiment below
removes only short duplicate projection walks while retaining the same
validated attribute walk; it does not recycle that candidate's hypothesis.

## One conditional next experiment

If the fresh 0811 native diagnosis shows that these short projection edges
matter, measure a single internal **per-slide capture projection** in
`SlidePart::finish_from_processed`. Its input is the existing MCE-processed
borrowed slice. Its bounded result would contain the validated `sld` root
projection, the producer-visible `cSld@name` result, and the same optional
`SlideRootProof` conformance. The implementation should fuse the two short
projections into the already-required complete notes scan, use one direct
`Reader<&[u8]>` plus `NamespaceResolver` state machine for the successful
common path, retain no DOM or second processed buffer, and leave the public
API, `SlideRootMemo`, and notes graph owner unchanged.

This experiment removes short duplicate walks; it does **not** remove the
100 complete notes-proof scans measured above. The profile makes its likely
benefit uncertain, so it should be abandoned before implementation if native
stacks show the short readers are immaterial. The lower-scale presentation
root/inventory fusion is a separate candidate and should not be bundled here.

This is a projection fusion experiment, not a validation weakening. The
state machine must defer any complete-proof refusal until root/name projection
has reached the same decision point as today, so a malformed name still wins
over a later notes-grammar error. It must preserve root-error masking and the
Strict/Transitional behavior; if a conformance retry is needed for an invalid
or adversarial input, an explicitly bounded retry is preferable to changing a
refusal. The retained notes graph remains authoritative and continues to
consume the proof by exact allocation identity.

Required evidence before retaining the candidate:

* A test-only differential oracle runs the old composition
  (`root_name_from_xml`, `presentation_name`, and
  `root_conformance_from_processed`) beside the fused projection. It compares
  names, fallback behavior, proof conformance, proof presence, and exact
  `Error` variant/message and precedence, including Transitional and Strict
  roots.
* The corpus covers no-notes and notes-bearing packages, MCE choices and
  unknown extensions, default/prefixed namespace rebinding, empty elements,
  malformed UTF-8, duplicate/reserved attributes, unknown prefixes, CDATA,
  processing instructions, DTDs, unterminated/multiple roots, and every
  depth/node/attribute/attribute-byte limit. Include missing, empty, malformed,
  and late `cSld@name` cases so name/refusal ordering is observable.
* Existing opened-capture, notes graph, slide-root memo, exact no-op, changed
  slide, allocation-identity miss, rebind, mixed-conformance, orphan-graph,
  preservation, and semantic readback tests pass. Add an assertion that the
  fused path retains no processed/XML allocation beyond the existing bounded
  proof/memo ownership.
* The frozen public workflow is recaptured for tiny, medium, and large
  attribute-free, vendor, unicode-vendor, and notes-bearing shapes across
  capture, commit, and lifecycle. Record native p50/p95/p99, operation
  allocation calls/bytes/net-live/peak, exact output bytes, semantic text,
  and refusal identities. Add an owner-scoped native profile to confirm that
  the removed reader passes, rather than an unrelated counter, explain any
  gain. Apply the existing benefit threshold and regression veto; an
  allocation-only reduction is insufficient.

If the differential oracle cannot preserve the current error precedence, or
if fresh native workflow evidence does not show a useful end-to-end benefit,
discard the candidate. Do not replace the proof with a relationship-based
no-notes shortcut and do not fold the lower-scale presentation scan into this
experiment.

## Source and evidence receipts

* [`docs/GOAL.md`](../../../../docs/GOAL.md) and accepted ADRs 0001, 0005,
  0006, 0010, 0011, 0013, 0030, 0031, and 0032 define the optimization order,
  lossless/refusal/resource contracts, lazy part ownership, and memo rules.
* [`change-0810/profile-analysis.json`](../change-0810/profile-analysis.json)
  and [`change-0810/profile-review.md`](../change-0810/profile-review.md)
  provide the measured profile boundary and direct-reader mechanism counts.
* [`change-0810/results-review.md`](../change-0810/results-review.md) records
  the retained 0810 workflow and allocation evidence.
* [`change-0806/results-review.md`](../change-0806/results-review.md) records
  the rejected iterator candidate and its frozen adoption gate.
