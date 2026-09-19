# Change 0693 code review — capture-local notes root proof reuse

Review status after the focused helper fix: **implementation accepted; all
correctness/build gates and source-bound measurements pass; trace/audit seal
remains parent-owned**.

The review covers the candidate diff against the 0693 baseline head
`5e89851b9e9d53daec5b291672d7ac15503243c7`, the 0693 design packet, `docs/GOAL.md`,
and the accepted validation, memory, snapshot, and physical-ownership ADRs.
This review is read-only with respect to Rust production code. The parent lane
recorded passing formatter, `cargo check`, Clippy, default, all-features,
facade, and rustdoc gates in `integration/results.json`, then refreshed the
native, allocator, profile, and refusal receipts against the same corrected
source. The default and all-features suites include the corrected private
source fallback and the swapped-sentinel notes-load proof test.

## What is sound in the candidate

* [`processed_xml_with_source`](../../../../crates/litchi-pptx/src/parts/mod.rs#L35)
  makes the deliberate two `Part::blob()` observations: the first is the
  existing 64 MiB preflight and the second is the slice passed to MCE. It does
  not call `process_part`, so the proof source and the MCE input are tied to
  the second observation and there is no hidden third observation.
* [`root_conformance_from_processed`](../../../../crates/litchi-pptx/src/notes/codec.rs#L320)
  runs the full existing scanner over the already processed bytes. It retains
  the raw and processed byte checks, root and namespace checks, depth, node,
  attribute, and relationship collection behavior, and keeps the
  Transitional-then-Strict order. Its `Option` result correctly masks scanner
  refusal details so a capture-time advisory scan cannot replace the later
  legacy notes error.
* [`capture_slides`](../../../../crates/litchi-pptx/src/presentation/package.rs#L49)
  reserves the optional proof vector fallibly and ignores a reservation
  failure, preserving the original capture path. A scanner `None` is retained
  as one proof entry and closes the prefix; later slides are still captured for
  the existing root, relationship, identity, and deferred-name ordering.
* The name result is evaluated before the notes proof, the processed `Cow` is
  dropped before a missing-name fallback allocates, and the capture entries are
  consumed before notes loading. The proof vector is explicitly dropped after
  the private notes snapshot call. It contains no processed XML, `Cow`, scanner
  inventory, or owned text. The borrowed `&[u8]` witness is intentional: only
  its pointer and length are compared, while the borrow keeps the exact second
  source allocation alive for the notes call and prevents ABA reuse. A slice
  plus `Option<Conformance>` entry is expected to be 24 bytes on the 64-bit
  target; the 0692 ABI receipt gives the corresponding layout assumptions.
* [`load_index_with_slide_root_proofs`](../../../../crates/litchi-pptx/src/notes/package.rs#L213)
  keeps content-type, relationship, notes-master/theme, notes-slide, orphan,
  aggregate-limit, materialization, and snapshot publication checks in their
  existing order. It checks the notes inventory length before enabling
  positional hints, checks the current slide blob once, and falls back to the
  old `root_conformance` path when the proof is absent or its source identity
  does not match. The conformance comparison with the presentation remains in
  place.
* The private capture-only snapshot wrapper leaves the public presentation and
  notes wrappers unchanged. The source is treated as immutable for the
  capture lifetime, consistent with the package contract; a changed pointer or
  length takes the legacy path and the proof never owns package bytes.

## Findings requiring disposition

### 1. Identity fallback is now enforced inside the helper

[`SlideRootProof`](../../../../crates/litchi-pptx/src/notes/package.rs#L63)
intentionally stores the exact second source slice as a borrowed witness. The
implementation reads only its pointer and length for identity; retaining the
borrow keeps the source allocation alive for the whole notes call and gives the
ABA/lifetime guarantee required by the design.

The earlier frozen private test passed `Some(&invalid_proof)` directly to
`resolve_slide_root` for a different allocation, a Strict allocation, and a
length-mismatched slice while expecting the legacy scanner fallback
([`notes/package.rs:1437`](../../../../crates/litchi-pptx/src/notes/package.rs#L1437)).
The helper now performs the `same_raw_source` filter itself, so both the
production caller and the direct unit test use the legacy scanner on a
mismatched pointer or length. The earlier runtime failure is resolved without
changing the production decision.

### 2. The advisory scanner still uses infallible vector growth

`scan_processed_xml` preserves the old `Vec::push` behavior for relationship
attributes and slide/master IDs. The new path invokes that scanner before later
capture refusals, so allocator exhaustion in those pushes can abort during a
proof scan where the old order would have returned a later root, relationship,
or name refusal first. Parser and finite-limit failures are correctly masked,
but the design's statement that scanner/allocation failure is advisory is not
fully true for allocator exhaustion. Either document this inherited abort
boundary explicitly and exclude it from the ordering guarantee, or make the
new proof-only collection fallible and treat reservation failure as `None`.

### 3. The documented scratch bound is stronger than the arbitrary `Part` path

`capture_internal` checks `limits.max_parts` against its first
`PresentationPart::slide_references()` result, while `capture_slides` parses
the catalog again and reserves from that second result. An immutable built-in
package produces the same count, but a foreign `Part` with changing catalog
views can make the second result reach the existing `MAX_SLIDES` bound of
100,000 after the first result passed a smaller policy. The candidate then
adds a proof reservation of roughly 24 bytes per second-catalog entry alongside
the already existing capture vector. This is finite and inherited from the
double catalog read, but the design/README wording that both reservations are
bounded by the already checked `max_parts` catalog should be narrowed to the
actual second-read `MAX_SLIDES` bound or the proof reservation should be capped
by the first checked policy.

### 4. Inventory-length fallback is intentionally conservative enough for the source contract

The notes hint gate compares only
`expected_catalog_len == presentation_scan.slide_ids.len()`. The later raw
pointer/length check catches a reordered or retargeted built-in slide. Exact raw
source identity is the relevant proof key: if a reordered entry resolves to
different bytes, the pointer or length check falls back; if immutable parts
share the same raw allocation, the already-proven classification is equivalent.
An equal-length relationship-ID comparison is therefore unnecessary under this
source contract. A swapped-sentinel integration test is still useful as a
regression demonstrating the fallback, but its absence is test coverage debt,
not a correctness blocker.

The source-changing foreign-part boundary also needs to remain explicit: a
pointer/length mismatch correctly falls back, while in-place mutation that
keeps both values unchanged is undetectable and therefore relies on the
immutable source contract. The review does not treat that contract-preserving
case as a new production bug, but it should be stated beside the witness
method.

## Ordering, limits, and memory checklist

The candidate preserves the intended ordering for ordinary immutable parts:
main root/catalog, per-slide relationship and root/name projection, all later
identity checks and deferred name replay, then notes validation. A proof `None`
or notes raw/processed limit is advisory during capture and the later legacy
notes scan retains its generic root refusal. Transitional/Strict classification
is masked in the same order, and a matching proof still compares with the
presentation conformance. Notes-master/theme/resource scans and outbound
relationship rejection are not bypassed.

The reservation policy is correct for the optional proof vector: an initial
`try_reserve_exact` failure disables hints and does not create an early
allocation error, and the successful reservation is large enough for the
catalog prefix so no later proof-vector growth is needed. The additional live
state is one 24-byte borrowed-slice proof entry per second-read catalog item;
processed slide bytes and scanner vectors die at each iteration. Names are moved into the
existing capture/final slide state rather than duplicated. The only remaining
transient overlap at notes load is the proof vector itself, which is dropped
after the private snapshot call.

## Tests and performance evidence

The added public matrix covers Transitional/Strict and MCE slides, mixed
conformance, late root/name/notes precedence, a foreign extra inventory ID,
the 16 MiB notes-root versus 64 MiB part limit, malformed tails, and separate
immutable package sources. The private source-identity test now exercises the
helper fallback. The added notes-load swapped-sentinel test covers equal-length
positional permutations: reordered invalid proofs fall back to the current raw
root, while exact identity retains the cached refusal. A first-name/proof-prefix
boundary with a later relationship refusal and a truly changing foreign
`Part::blob()` source remain useful future coverage. The recorded default suite
passes 918 tests with two ignored; all-features passes 932 with two ignored;
the facade suite passes 45 tests with no failures.

The expanded refusal ABBA measurement is now available. The candidate adds
about 44--49 microseconds at p50 for the two-slide late-root and late-missing-
relationship refusals: marker-free cases rise from about 35.0/27.8 microseconds
to 79.1--79.5/71.7--72.1 microseconds (roughly +123--160%), while MCE-bearing
cases rise from about 102.7--113.2 to 152.1--160.1 microseconds (roughly
+41--48%). Typed outcomes and input hashes agree. This is the expected cost of
relocating two complete earlier-slide scans before a later refusal, and it is a
real refusal-path cost rather than measurement noise.

The final source-bound native rerun materially improves common successful
marker-free workflows. Total medians improve by 29.42%/29.11% for real one,
34.22%/34.68% for real no-op, and 28.83%/28.93% for real two-slide workflows;
marker-control improves 8.8--16.7%, and generated inputs improve 6.9--17.1%.
The notes workflows stay within -2.98% to +2.32% at total p50. No total-phase
median, mean, p95, or p99 exceeds the 5% trigger. The report retains 28
non-total phase-tail triggers (7 p95 and 19 p99, plus one p50 and one mean):
the clearest is the no-op LibreOffice notes case's apply phase at +14.46%
(+2.32 microseconds) p50 and +14.71% (+2.39 microseconds) mean. The no-op
generated pair has an elevated baseline leg (0.9082 ms versus 0.8301 ms), so
that result remains tied to its leg metadata. Allocation diagnostics show a
substantial drop in calls and requested bytes, with peak live bytes nearly
flat.

The refreshed prefix profile also improves per-open-capture cycles by 27.57%,
instructions by 31.70%, branches by 32.02%, branch misses by 17.27%, cache
misses by 12.80%, page faults by 17.15%, and task-clock by 27.64%; whole-child
peak RSS rises by 92 KiB (5540 to 5632 KiB). The counting allocator's peak-live
and net-live values remain flat or differ by only the measured few KiB while
calls and requested bytes fall substantially.

Disposition: retain the broad marker-free reuse opportunity rather than gating
only to MCE-owned/`Cow::Owned` slides; that gate would discard the measured
common gains. Treat the refusal amplification as an explicit boundedness cost,
not as an unqualified win. The owner accepts the measured +44--49 microsecond
late-refusal cost under the unchanged finite slide/XML scanner bounds in
exchange for the common real/control/generated gains; no extra proof-work
budget or MCE-ownership gate is required for this change. The final packet must
include the complete native, allocator, profile, and correctness gates; the
README remains `performance_claim: none`, including after sealing.

## Review disposition

The focused correctness matrix, including the swapped-sentinel notes-load test,
passes in the recorded default and all-features suites. The borrowed-slice
witness and the length-only inventory gate are accepted design decisions. The
infallible scanner behavior must remain documented as an abort/resource
boundary rather than being described as a typed advisory allocation failure.
The final source-bound native, allocator, profile, and refusal rerun is
recorded; only the parent-owned trace restoration, audit, and packet seal
remain outside this code review.
