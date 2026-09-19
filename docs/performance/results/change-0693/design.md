# Change 0693 design: reuse capture-local notes root proofs

Status: retained after the resolver identity-test fix, refreshed measurements,
all quality/repository gates, trace restoration, and code review. The baseline
is `5e89851b9e9d53daec5b291672d7ac15503243c7`; source, build-input and accepted
constraint hashes are retained in [`baseline.json`](baseline.json).
`performance_claim: none` remains after validation and sealing.

The candidate is a narrow extension of the capture-local slide projection. It
reuses a conformance proof for the later notes validation without retaining
processed XML, changing a public view, or making the notes graph lazy.

## Current path and repeated work

`opened::capture_internal` first validates the PresentationML main part,
parses its slide catalog, constructs the borrowed presentation view, captures
each slide, checks slide relationships and identities, and only then calls
`notes::load_snapshot` ([`opened/model.rs:612`](../../../../crates/litchi-pptx/src/opened/model.rs#L612)).
The current capture helper
[`SlidePart::from_part_with_name`](../../../../crates/litchi-pptx/src/parts/slide.rs#L902)
does one `processed_xml` call, validates the `p:sld` root, reads the
producer-visible `p:cSld@name`, drops the `Cow`, and returns the borrowed
`SlidePart` plus the deferred name result. A name error is held at its original
slide index while subsequent slide roots are still checked; the capture loop
replays relationship, target, identity, and name ordering before returning it.

The notes path then repeats work over a slide. `load_index` first determines
the presentation conformance and scans the presentation inventory. For every
referenced slide it calls `root_conformance` with the notes slide limit. That
helper tries Transitional and Strict and runs the complete notes XML scanner
for each attempt, discarding failed-attempt detail. The scanner itself calls
MCE before applying the root, depth, node, attribute, and byte policies. The
same notes path still has separate resource scans for the notes master, theme,
and notes-slide XML. Those resource scans and their relationship-attribute
refusals are outside this proof.

## Capture proof

The low-level preprocessing helper at
[`parts/mod.rs:26`](../../../../crates/litchi-pptx/src/parts/mod.rs#L26) must
keep two deliberate `Part::blob()` observations:

1. The first observation supplies the existing `MAX_PART_XML_BYTES` check.
2. The second observation is bound to the slice passed to MCE. Its pointer and
   length are the only raw identity witness that may be carried forward.

The helper must not reuse the first return value for both purposes, and it
must not call a wrapper such as the current `process_part(part)` if that
wrapper obtains a third `blob()` view internally. The MCE call, the returned
`Cow`, and the witness all refer to the second observation. Existing MCE input
and output limits remain in force.

The capture-only slide helper then performs the following sequence for each
slide while the one processed `Cow` is live:

1. Validate the slide content type and `p:sld` root exactly as today.
2. Evaluate the producer name over that same processed slice. A successful
   missing name still becomes the existing part-name fallback.
3. Only after the name result succeeds, run the complete notes XML scanner on
   the already processed slice. This private scanner form must not call MCE a
   second time. It must apply the notes slide raw and processed byte limits,
   root and namespace rules, XML depth/node/attribute limits, and attribute
   collection behavior that the later `root_conformance` call would apply.
4. Drop the processed slice and all scanner inventory. Retain only a bounded
   proof entry containing the exact second borrowed raw-slice witness (checked
   by pointer and length) plus an `Option<Conformance>`.

`Some(Conformance)` means the already processed slice passed the notes scan in
the normal Transitional-then-Strict order. The first invalid scanner result is
retained as an entry with `None` conformance and closes the prefix. If its raw
witness still matches at notes load, return the same generic
`Invalid("invalid sld root or namespace")` result directly; do not re-run MCE.
A mismatch at that invalid entry is treated like any other identity mismatch
and uses the legacy scanner. A normal slide root failure remains immediate as
it is today. On the first name error, the proof prefix ends before that slide,
while the existing capture code continues all required later root checks and
keeps the original deferred name error. No proof is collected after either
boundary.

The proof is deliberately a prefix in the original PresentationML catalog
order, rather than a compact list of notes-bearing slides. This makes a proof
for slide `i` usable only at slide `i` and leaves entries after a first name or
proof failure untrusted. The proof vector is separate from the capture-entry
vector. On the measurement target its entry is expected to be about 24 bytes
per slide; it contains no XML, `Cow`, `Arc`, `XmlScan`, or owned text. Capture
entries are about 48 bytes per slide and should be released before entering
notes loading when the ownership flow allows it. For `N` equal to the actual
capture-catalog bound, the two vectors account for approximately `24N + 48N`
bytes before allocator overhead. An ordinary immutable input under the default
`max_parts=4096` therefore has the qualified ~294,912-byte two-vector bound;
that is not a universal bound. The low-level `MAX_SLIDES=100000` guard,
caller-custom limits, and a foreign `Part` whose catalog changes between the
initial references and capture catalog all require the actual `N` and existing
mismatch handling to remain authoritative. If the optional proof vector cannot
reserve, disable hints and continue the root/capture path; do not turn that
scratch allocation failure into a new early document error.

## Notes hint and fallback

Add a private `load_snapshot` variant for the capture caller. Existing
transaction and public-facade callers continue using the legacy variant. The
capture variant receives the expected catalog length and the proof prefix; it
does not expose either as API state.

The notes presentation scan produces the slide-ID inventory that currently
drives `load_index`. Before using any proof, compare that inventory length
with the expected catalog length. If they differ, disable every hint for this
load and run the legacy checks. This guard is global because positional proof
alignment is no longer established. The graph must still be fully validated
and materialized in the old order.

For every presentation slide, preserve its original catalog position while
following the notes relationship. The loop still checks slides with no speaker
notes because slide-root conformance precedes the optional notes relationship.
At that position, a hint is usable only when all of the
following hold:

* hinting was not disabled by the inventory-length check;
* the proof entry exists;
* the current slide part's `blob()` is obtained once, and its pointer and
  length exactly match the proof's second-observation witness.

The current slice is an identity check only; it is never retained or used as a
replacement for the proof's processed XML. If any condition fails, pass that
single current slice through the legacy `root_conformance` behavior. If the
entry matches but contains `None`, return the existing generic invalid-root
error directly without another MCE call. Do not use a pointer alone,
`blob_arc()` identity, or a compact notes-only index.

When a hint matches, compare its `Conformance` with the presentation
conformance exactly as the old root check did. A mismatch still returns the
existing typed invalid result. A matching hint replaces only the repeated
slide root/conformance scan. It does not skip slide content-type checks,
relationship target checks, notes-master/theme checks, notes-slide resource
scans, orphan detection, `Snapshot::from_parts`, or any graph materialization.

The existing `validate_resource_xml` path must continue scanning the notes
master, theme, and notes-slide payloads and then rejecting outbound
relationship attributes. The proof scanner may discard its own inventory
after recording conformance, but it must not turn those later resource checks
into hints.

## Ordering and resource semantics

The candidate must preserve the following observable sequence:

* main-part content type/root and catalog validation happen before slide
  capture;
* every capture relationship, target, identity, root, and deferred-name check
  retains its current position;
* notes loading starts only after the capture identity loop succeeds;
* `load_index` performs presentation inventory, notes-master cardinality,
  slide relationship/root checks, master/theme checks, notes relationships,
  orphan checks, and limits in the current order;
* graph materialization and source-checked snapshot construction happen after
  indexing, with no partial publication.

The proof scanner is advisory for typed document and scanner-limit errors. It
cannot replace a pre-existing earlier capture refusal with a different typed
error; failure to reserve the optional proof vector falls back to the old
capture path. Infallible growth inside the proof-only relationship and ID
collections remains an explicit allocation-abort boundary. If the capture
reaches notes loading, a missing proof or identity mismatch follows the old
scanner; a matching `None` entry returns the same generic invalid-root error
without re-running MCE. A proof hit can remove a later MCE rewrite and scanner
pass only because the same raw slice was processed and fully scanned after the
successful name.

This changes refusal work in a visible way. A late slide-root refusal can now
follow full proof scans for earlier successful slides. The capture loop checks
a slide's relationship before preprocessing that slide, so a late relationship
refusal can include completed proof scans only for earlier slides. The corrected 0693
refusal baseline exercises these paths; its `notes-invalid-tail` case retains
the direct XML refusal from the notes resource scanner. The initial malformed
fixture serialization attempt is retained as setup history and is not timing
evidence.

The candidate also removes an independent MCE allocation-failure opportunity:
the old notes root check could allocate a second MCE output after capture had
already succeeded, whereas the candidate uses the already processed bytes and
only retains a small proof. This is a resource difference, not permission to
change typed document errors or allocation limits. The proof-only scanner keeps
the existing infallible vector growth for collected relationships and IDs, so
allocator exhaustion during that scanner can still abort before a later
refusal. Only failure to reserve the optional proof vector is recoverable and
disables hints. Allocation diagnostics and refusal measurements must report
both the removed repeated-MCE allocation opportunity and this remaining abort
boundary explicitly.

## Fresh candidate evidence

The post-fix native A/B/B/A run retains 78 process records and 7,800 timed
samples across the 13 cases. Total-phase p50 changes are paired by `b0/a2` and
`b1/a3`: real one-edit is −29.42%/−29.11%, real no-op −34.22%/−34.68%, and
real two-slide −28.83%/−28.93%. Control changes are −9.57%/−10.78% for one,
−16.71%/−16.03% for no-op, and −8.77%/−11.15% for two. Generated changes are
−8.82%/−8.99%, −17.08%/−9.57%, and −6.88%/−7.77% for the same three
workflows. The two notes corpora range from −2.98% to +2.32% across their one
and no-op totals; no total-phase p50 is more than 5% slower. The generated
no-op baseline legs are `0.9082` and `0.8301` ms, so the paired values are
retained rather than averaged. The one median/mean review trigger is a
no-op-notes-Lo apply phase (+14.46%/+14.71%, about +2.32/+2.39 µs); it does not
represent a total-phase regression.

The separate allocator diagnostics for one-real change allocation calls from
270,242 to 190,338, reallocations from 6,950 to 4,870, and requested bytes
from 19,915,826 to 13,815,905. Peak live bytes are 463,159 versus 459,170 and
net live bytes remain 195,730. Prefix profiling changes per open-plus-capture
from 39.738M to 28.782M cycles, 159.173M to 108.711M instructions, and 211.32
to 175.075 page faults; whole-child peak RSS is 5,540 versus 5,632 KiB. These
are diagnostics and do not replace phase timing.

The ten-case refusal table in [`refusal-tables.md`](refusal-tables.md) shows
late-root and late-missing-relationship p50 increases of 43.95–48.72 µs,
125–161% for marker-free cases, and 42–47% for MCE-bearing cases. The added
work is the complete proof scan for earlier successfully named slides. The
relationship check for a late slide still precedes preprocessing that slide,
so its added work covers earlier slides only. Typed outcomes and input hashes
remain equal, and the removed repeated MCE allocation-failure opportunity is a
separate resource difference.

## Verification obligations

Focused tests should cover:

1. marker-free, MCE-bearing, Transitional, and Strict slides, including a
   missing name, a name parser error, and malformed XML after an otherwise
   readable name;
2. proof-prefix termination at a scanner `None` and at the first name error,
   with later roots and required identity checks still exercised in their old
   order;
3. a matching witness, a pointer/length mismatch, a missing proof, and a
   notes inventory length mismatch, proving that each fallback is complete;
4. notes slides that occupy nonconsecutive catalog positions, proving hints
   use the original position rather than the compact notes-slide order;
5. mixed slide conformance, duplicate names, late invalid slide roots, late
   missing relationships, invalid notes tails, notes-size limits, and orphan
   graph cases, preserving the old typed error family and precedence;
6. no-op revision/output identity, two distinct edited slides with exactly two
   excluded payloads, order-insensitive metadata oracles, and full reopen
   preservation.

The native 0693 packet contains a 13-case A/A baseline and a source-bound
A/B/B/A comparison; each process uses five warmups and 100 samples. The
allocator baseline and candidate remain separate 13-process diagnostics, and
the corrected refusal matrix has ten cases with two baseline and two candidate
legs. The exact-baseline rebuild and restoration are retained in
[`expanded-refusal-baseline/manifest.json`](expanded-refusal-baseline/manifest.json);
the earlier eight-case material is retained under `refusal-eight-case/`. The
prior 0692 trace is reused only as a source-equality-bound diagnostic call-count
baseline, recorded in [`baseline-trace-reuse.json`](baseline-trace-reuse.json);
it supplies no latency number. The current native, allocator, profile, refusal,
report, and repository-gate receipts are bound to the corrected source. The
trace and final evidence audit/seal remain before final retain. Profiles and
MCE/call traces may explain a result, but they cannot replace the native phase
evidence. No performance claim is made by this design packet.

The reproducible driver sequence is:

```text
# At baseline_head, before applying the candidate source:
python3 docs/performance/results/change-0693/build.py baseline
python3 docs/performance/results/change-0693/measure.py baseline
python3 docs/performance/results/change-0693/build-refusal.py baseline
python3 docs/performance/results/change-0693/measure-refusal.py baseline
python3 docs/performance/results/change-0693/measure-allocations.py baseline
python3 docs/performance/results/change-0693/profile.py baseline

# With the candidate source restored:
python3 docs/performance/results/change-0693/expanded-refusal-baseline.py
python3 docs/performance/results/change-0693/build.py candidate
python3 docs/performance/results/change-0693/measure.py compare
python3 docs/performance/results/change-0693/build-refusal.py candidate
python3 docs/performance/results/change-0693/measure-refusal.py compare
python3 docs/performance/results/change-0693/measure-allocations.py candidate
python3 docs/performance/results/change-0693/profile.py candidate
python3 docs/performance/results/change-0693/report-metrics.py
python3 docs/performance/results/change-0693/trace.py --profile candidate --model-trace --with-generated --with-control --cpu 12 --target-dir ../litchi-target-0693
python3 docs/performance/results/change-0693/trace-summary.py
python3 docs/performance/results/change-0693/run-integration.py
python3 docs/performance/results/change-0693/quality-summary.py
python3 docs/performance/results/change-0693/audit.py
python3 docs/performance/results/change-0693/seal.py
```

The trace remains diagnostic and uses the same sibling target directory as the
other probe drivers. The seven repository quality commands record exit 0, with
918 default-test passes and 932 all-feature passes (two existing ignored tests
in each suite) plus 45 facade passes. The diagnostic trace receipt passes with
no skipped runs; the final audit/seal commands remain verification steps and do
not imply final retain in this packet. The baseline build commands are
source-bound to
`baseline_head`; `expanded-refusal-baseline.py` is the candidate-state bridge
that temporarily restores the baseline source, builds the ten-case refusal
binary, and restores candidate bytes exactly.

## ADR compliance

The proposed seam is constrained by the accepted repository decisions:

* [ADR 0003](../../../../docs/adr/0003-snapshots-edits-and-patches.md): the
  proof is capture-local scratch for an immutable source and does not alter
  snapshot publication, edit, patch, or conflict state.
* [ADR 0005](../../../../docs/adr/0005-io-memory-and-performance.md): the
  optional proof vector is bounded and releasable; processed XML is not a
  persistent cache, and reserve failure disables the hint rather than adding
  an early refusal.
* [ADR 0006](../../../../docs/adr/0006-validation-security-and-compatibility.md):
  full scanner limits, graph checks, typed document errors, and error order
  remain in force; a hint only removes duplicate work already proven on the
  same source identity.
* [ADR 0011](../../../../docs/adr/0011-ooxml-physical-package-ownership.md):
  OPC parts and their bytes remain owned by `litchi-opc`; the PPTX layer holds
  no physical-package replacement or retained XML allocation.
* [ADR 0024](../../../../docs/adr/0024-current-topology.md): the change stays
  within the existing PPTX/notes/common/OPC dependency direction and adds no
  workspace topology edge.
