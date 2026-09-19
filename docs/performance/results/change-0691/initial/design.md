# Change 0691 design: fuse the PPTX notes root scan with its first complete scan

Status: design only. No production source, test, benchmark, or public API change
is included in this note. The fresh baseline is the current branch head
`bd50cb0d554d9adbc6ca772a279ba36973a0dd21`; the packet records the source and
accepted-constraint hashes in [`baseline.json`](baseline.json).

This review covers the opened PPTX capture path and the shared MCE calls it
reaches. It excludes iWork and does not propose changing the MCE namespace
emission contract from change 0653/0666.

## What the current capture does

`opened::capture_internal` first validates the PresentationML main part, parses
the slide catalog, creates a borrowed `Presentation` view, resolves every slide,
and then asks each resolved slide for its name before it loads the complete
notes graph. The sequence is visible at
[`opened/model.rs:618`](../../../../crates/litchi-pptx/src/opened/model.rs#L618):

1. `PresentationPart::from_package` calls `PresentationPart::from_part`, which
   runs `processed_xml` once to validate the main root.
2. `presentation.slide_references()` runs `processed_xml` over the same main
   part again.
3. `Presentation::slides()` calls its existing catalog memo, whose first fill
   parses `slide_references()` a third time, then constructs each `SlidePart`.
   Each `SlidePart::from_part` runs `root_name` and therefore one MCE pass over
   that slide.
4. The capture loop calls `slide.name()`. `c_sld_name` processes that same
   slide again, so the root and name projections are two independent passes.
5. `notes::load_snapshot` validates the notes graph. Its `load_index` first
   calls `root_conformance` for the presentation and then calls `scan_xml` again
   to obtain the `XmlScan` it immediately needs. On a valid Transitional part
   this is two MCE calls; on a valid Strict part the discarded Transitional
   attempt makes it three. Notes master, theme, and each notes slide are each
   scanned once by `validate_resource_xml`.

The shared helper is deliberately small:
[`parts/mod.rs:26`](../../../../crates/litchi-pptx/src/parts/mod.rs#L26) checks
the 64 MiB PresentationML part ceiling and calls
`litchi_ooxml_common::mce::process_part`; `process_part` is the plain
`process_ooxml(part.blob())` route at
[`mce/codec.rs:1427`](../../../../crates/litchi-ooxml-common/src/mce/codec.rs#L1427).
On a marker-free part this route performs the bounded marker scan and input
limit check, then returns a borrowed `Cow`; it does not itself parse XML or
perform a full MCE rewrite. The downstream root/name readers still parse their
own projections. Marker-bearing parts take the owned MCE rewrite path.

The notes duplication is especially cleanly bounded. At
[`notes/package.rs:166`](../../../../crates/litchi-pptx/src/notes/package.rs#L166),
`root_conformance` is followed immediately by the second `scan_xml` at line
167. `root_conformance` itself is only a loop over the two conformance choices
at [`notes/codec.rs:311`](../../../../crates/litchi-pptx/src/notes/codec.rs#L311);
it discards every failed scan error and returns the same generic
`invalid {root} root or namespace` error when neither choice succeeds.

## The recommended first change

Replace the notes pair with one helper that returns both the selected
conformance and the successful `XmlScan`:

```rust
pub(crate) fn root_conformance_and_scan(
    xml: &[u8],
    max: usize,
    root: &str,
) -> Result<(Conformance, XmlScan)> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        if let Ok(scan) = scan_xml(xml, max, conformance, root) {
            return Ok((conformance, scan));
        }
    }
    Err(invalid(format!("invalid {root} root or namespace")))
}
```

`load_index` then becomes one destructuring call:

```rust
let (conformance, presentation_scan) =
    root_conformance_and_scan(presentation.blob(), MAX_PRESENTATION_XML, "presentation")?;
```

This is pass fusion, not a cache. It keeps the first successful bounded scan
that the caller already paid for; it does not retain processed XML after
`load_index` returns, add an `Arc`, or alter any part type. The failed
Transitional attempt for a Strict input remains exactly where it was. The
successful scan's `XmlScan` is the same value that the old immediate second
`scan_xml` call would produce from the same immutable bytes.

The change removes one successful MCE rewrite and one complete XML validation
pass for a valid Transitional notes graph. A valid Strict graph still performs
the failed Transitional probe and then removes the duplicate successful Strict
pass. It also removes the duplicate traversal of relationship attributes and
the associated temporary vectors. The gain is intentionally left unmeasured
until the root agent's frozen baseline lane is available.

The proposed helper must stay private to the notes codec/package boundary. It
does not generalize `scan_xml` into a public parser, and it does not change
`validate_resource_xml`: that function must continue to call `scan_xml` and only
then reject nonempty `relationship_attributes` at
[`notes/codec.rs:295`](../../../../crates/litchi-pptx/src/notes/codec.rs#L295).

## Error order, limits, and graph ownership

The existing order is load-bearing:

* `load_index` checks the presentation content type, then root/conformance and
  the complete presentation scan, before inspecting notes-master cardinality,
  slide IDs, slide relationships, slide roots, master/theme topology, notes
  relationships, and orphan parts.
* `scan_xml` checks raw input bytes and processed output bytes before parsing;
  it then enforces `MAX_DEPTH`, `MAX_NODES`, `MAX_ATTRIBUTES`, and
  `MAX_ATTRIBUTE_BYTES`, checks the expected root and namespace, and collects
  relationship attributes. These checks are implemented at
  [`notes/codec.rs:320`](../../../../crates/litchi-pptx/src/notes/codec.rs#L320).
* A failed root-conformance probe currently throws away the detailed scan error
  and returns the generic invalid-root error. The proposed helper must retain
  that behavior. It must not return a parser, MCE, allocation, or limit error
  from a failed probe where the old `root_conformance` returned the generic
  error.
* Once one conformance scan succeeds, the old second scan has the same document
  inputs and policy, although a separate run could still encounter a transient
  allocation failure. Reusing the successful `XmlScan` therefore preserves
  document-level semantics and error precedence while deliberately reducing
  that repeat allocation opportunity; it moves no content or policy refusal
  and skips no validation. `Snapshot::from_parts` and its bounded
  relationship/part checks remain after materialization in the same order.

The notes limits are finite and unchanged: presentation 32 MiB, slide 16 MiB,
notes slide 8 MiB, notes master/theme 16 MiB, aggregate notes bytes 64 MiB,
4,096 notes slides/owned parts, 100,000 nodes, depth 128, 500,000 attributes,
and 8 MiB of attribute bytes. The helper consumes the exact `XmlScan` produced
under those limits; it introduces no second output buffer and no new ceiling.

## Why a persistent processed-XML memo is not the first step

`Part::blob_arc` is useful evidence but is not a universal identity guarantee.
The trait only promises an `Arc<Vec<u8>>` view at
[`opc/part.rs:51`](../../../../crates/litchi-opc/src/part.rs#L51); a foreign
`Part` implementation may return an `Arc` whose allocation or contents do not
alias `blob()`. Built-in `BlobPart` and `XmlPart` do preserve the same
`PartPayload` allocation, and `OpcPackage::get_part` forces deferred payloads
before handing the part out at
[`opc/package.rs:599`](../../../../crates/litchi-opc/src/package.rs#L599).
Mutating a built-in part replaces its payload `Arc`, rather than mutating bytes
through an existing allocation.

The existing opened-presentation digest memo demonstrates the required proof:
`memo_key` checks equal length and pointer equality between `blob_arc()` and
`blob()`, then retains the `Arc` so an address cannot be recycled
([`opened/model.rs:413`](../../../../crates/litchi-pptx/src/opened/model.rs#L413)).
A processed-XML memo would need the same proof plus ownership of the processed
output. A pointer-only key is unsound for foreign parts and an unretained raw
address has an ABA hazard.

The borrowed `PresentationPart` and `SlidePart` views are public `Clone + Copy`
wrappers around `&dyn Part` ([`parts/presentation.rs:39`](../../../../crates/litchi-pptx/src/parts/presentation.rs#L39),
[`parts/slide.rs:860`](../../../../crates/litchi-pptx/src/parts/slide.rs#L860)).
Adding `OnceLock` state to either would break that public shape. `Presentation`
already owns a bounded catalog memo because its immutable package borrow makes
that state safe for the view's lifetime, but a processed MCE buffer can be as
large as the part and would remain live with the view.

The long-lived opened `Snapshot` owns an `Arc<OpcPackage>` and is cloned into
transactions. Retaining rewritten XML there would pin potentially large
allocations across edits and publications. ADR 0005 requires retained state
whose size is the point to have an explicit finite policy, observable ownership,
and a release path. The existing `Limits::max_retained_candidate_bytes` is
specifically the cross-slide candidate archive ceiling; its documentation at
[`opened/model.rs:138`](../../../../crates/litchi-pptx/src/opened/model.rs#L138)
must not be repurposed as an MCE cache budget. The notes pass fusion has none of
these retention costs.

## Preferred higher-impact follow-up: one MCE pass for each slide root and name

The notes fusion is the smallest proof, but the opened capture's slide root and
name projections account for the larger repeated work on an MCE-heavy deck.
Change 0649 measured 13 `SlidePart::from_part` root passes and 13
`SlidePart::name` passes per capture on its 13-slide real deck
([`0649-pptx-opened-transaction-real-deck-edit.md:181`](../../../../docs/performance/0649-pptx-opened-transaction-real-deck-edit.md#L181)).
The root pass happens while `Presentation::slides()` is building every slide;
the name pass happens only after all roots have succeeded, in the identity loop
at [`opened/model.rs:638`](../../../../crates/litchi-pptx/src/opened/model.rs#L638).
This is the next focused candidate if the fresh baseline confirms that MCE
rewrite cost still dominates notes validation.

The fresh 0691 baseline supports prioritizing this seam. The real-deck native
capture p50 is 10.660–10.754 ms versus 3.357–3.417 ms for the marker-free
control, and the edit-phase total is 22.821–23.056 ms versus
7.534–7.622 ms; all four legs are retained in
[`native-summary.json`](native-summary.json). The capture trace records 44
calls over 819,319 bytes and 14 raw identities, all owned. These are baseline
observations rather than an attribution or a speedup claim: they show that the
real route has a meaningful MCE-bearing capture cost, while the proposed tests
and an after trace must still establish exactly which passes disappear.

The safe seam is a capture-only, crate-private deferred projection, with no
processed-XML retention:

1. Factor the existing `root_name` and `c_sld_name` readers so they can accept
   an already processed byte slice. Add a crate-private `SlidePart` constructor
   used only by capture that checks the content type, calls `processed_xml(part)`
   once, validates the root immediately, and then evaluates
   `namespace::presentation_name` over that same `Cow`.
2. Return the validated `SlidePart` plus a deferred name result. A root error
   is returned immediately, exactly as `SlidePart::from_part` does. A name
   error is held until the capture identity loop; a successful producer name or
   the existing part-name fallback is the final `String` that the opened
   snapshot needs anyway.
3. Add a crate-private `Presentation`/package helper used only by
   `opened::capture_internal` to preserve the existing relationship/content
   checks and build the contextual slide vector. After that vector has finished
   all slide root checks, the existing loop must repeat the current
   relationship-target and identity checks in the same order, then consume the
   deferred name result at that slide's original position.

The important point is that the name result is deferred, not the root check.
For example, if slide 0 has a malformed `p:cSld` name projection and slide 1
has a bad root, the helper must return slide 1's root error. If slide 0 has a
name error and slide 1 has a duplicate identity, the existing loop must return
slide 0's name error before it examines slide 1. If slide 0 is valid and slide
1 has the duplicate identity, the duplicate must still beat slide 1's deferred
name error. The same ordering applies to the repeated relationship lookup and
target comparison in the existing capture loop. Notes loading remains after
all identities and names, unchanged.

The simplest implementation can carry `Vec<Result<String>>`, bounded by the
already enforced `limits.max_parts` check. A stricter transient-memory variant
should carry successful names and only the earliest name error (with its slide
index), because no later name error can become observable once that earlier
error is replayed. That avoids retaining one dynamically formatted error per
slide on a hostile package. Either form must use fallible vector reservation
and document the temporary result storage as capture scratch; it must not add a
field to public `Clone + Copy` `SlidePart`/`PresentationPart`, a snapshot cache,
or an interpretation of `max_retained_candidate_bytes`.

This seam is safe for foreign `Part` implementations without relying on
`blob_arc` identity: the processed `Cow` is consumed before the helper returns,
and no key or retained allocation is published. It does rely on the existing
capture invariant that a borrowed package's part bytes remain stable during one
operation; built-in deferred payloads are forced and memoized by
`OpcPackage::get_part`. A marker-free part still gets only the one bounded MCE
marker scan, while marker-bearing parts avoid the second rewrite allocation.
The two readers may remain separate over the shared processed slice so their
established namespace and early-return behavior is preserved; combining them
into a new parser would add unnecessary semantic risk.

The optimization can change only transient allocation behavior: the old second
`processed_xml` call could fail to allocate even after the first succeeded, and
the fused call removes that repeat opportunity. It must preserve all document
errors, typed limits, root/name precedence, relationship checks, and output
values. Focused tests should cover transitional/strict namespace forms, MCE
markers, missing/empty `cSld@name`, malformed tails, first/second-slide root and
name failures, duplicate identities, relationship target failures, and absent
or invalid notes graphs. Instrumentation should show one `process_ooxml` call
per slide in capture on valid inputs, with no changed call count for the main
part or notes graph.

This candidate is preferred over a persistent per-part processed-XML cache for
this batch: it removes the two large per-slide rewrites seen in the real-deck
profile while retaining only values capture already stores, and it leaves ADR
0005's retention budget and `blob_arc` ABA proof out of the production seam.

## The other repeated calls

The main presentation's third catalog parse can be removed later by adding a
crate-private unvalidated catalog accessor to `Presentation` and making capture
use that memo before calling `view.slides()`. Calling the public
`slide_references()` method is not equivalent: it runs whole-catalog
relationship validation and would change the existing distinction between
unvalidated `slides()` and validated `slide_count()`/`slide_references()`.
This is a small, safe follow-up, but its value is narrower than the notes and
slide root/name fusions.

## Proof obligations before production code

An implementation should add focused tests and run the existing PPTX/MCE gates
for:

1. valid Transitional and Strict presentation roots, with and without MCE
   markers, asserting the exact same `Conformance`, `XmlScan`-derived graph,
   notes snapshot revision, and published bytes;
2. wrong roots, missing roots, malformed XML, malformed MCE, unbound prefixes,
   and every input/output/node/depth/attribute limit, asserting the same typed
   error category and message as the old helper;
3. presentation and notes parts carrying outbound relationship attributes,
   proving `validate_resource_xml` still reports its relationship refusal after
   the complete scan;
4. absent notes graphs and orphan notes graphs, proving the same topology checks
   and no partial snapshot publication;
5. repeated notes capture, transaction commit, inverse application, and stale
   patch checks, proving the fused `XmlScan` is scoped to one `load_index` call
   and cannot outlive the immutable source it describes.

Instrumentation should show one fewer successful `scan_xml`/`process_ooxml`
call per notes `load_index` on the valid paths and no change to the calls on a
root-conformance refusal. The final before/after packet must retain the frozen
baseline, command identities, corpus hashes, and a stated A/A floor. No
performance claim follows from this design note alone.
