# XLSX profile validation and capture history

The current sealed baseline is [the `ab954d91d` capture](results/clean-ab954d91d/README.md).
It is approved for capture integrity and descriptive measurements, with no
causal speedup or regression claim. The following readiness notes retain the
history of earlier scaffolds and provisional evidence.

## Scaffold development history

This handoff contains requirements, runner infrastructure, and an adapter
wired to the current public XLSX selector / source scanner / attach / detach
transaction signatures. The production semantics and measurement freeze are
still under review, so there are no final or acceptance metrics or claims.
`results/exploratory-before/` retains a separately labeled bounded baseline for
the current working tree's 16/64/256 same-drawing attach lanes.
`results/exploratory-detach-before/` is the corresponding six-lane shared and
distinct detach probe, also exploratory-only.
`results/exploratory-namespace-before/` is the corresponding 32-picture,
distinct-target, shared-root-namespace inventory probe; its retained-source
byte observation is cross-referenced to the source-scope probe.

The adapter uses deterministic synthetic packages adapted from
`crates/litchi-xlsx/tests/drawing_svg_lifecycle.rs`, plus the retained native
LibreOffice fixture. It exercises all three anchor forms, shared and distinct
targets, inventory counts, namespace pressure, strict-host and incoming-edge
lifecycle cases, captured-owner clone cases, same-picture attach/detach,
two-worksheet attach/detach, and refusal lanes. The exploratory
root-namespace lane requires raw-source and shared-context projections for all
32 owners. It must
still satisfy every semantic and refusal gate in `requirements.md` before
setting `XLSX_SVG_PROFILE_API_WIRED=1`. The three dedicated multi-picture
attach lanes and six exploratory detach lanes stage 16, 64, and 256 edits on
one drawing and validate one composed commit/reopen per count. They are
bounded cost lanes for the current per-intent rescan behavior, without a
scaling claim. The 69-lane acceptance matrix adds exact in-memory inverse and
replay-refusal checks for every anchor/size pair and a composite mixed-cap
atomic-refusal lane. The adapter repeats OPC reachability, SVG leaf, content
type, fallback, and selected-owner checks after each changed publication;
source-copy and cumulative staged-byte fields stay null until the production
owner exposes those boundaries. Shared-final inputs are built with only the
selected picture owning the shared SVG, without production detach setup. The
incoming-edge lane adds a separate package relationship that must retain the
SVG leaf after final drawing-owner detach. Strict attach and detach lanes
assert the strict worksheet/drawing relationship types. Captured-owner clone
lanes scan and validate once before warm-up and measurement, then time only
cloning the captured `PictureSource`, separate from workbook open/save and
capture parsing. Namespace refusal derives the first refused active-binding
count by probing the admitted policy, and opaque extension/descendant checks
compare exact bytes and counts. Receipts distinguish generated pressure
bindings, total active bindings including seven fixed root declarations, and
the refused active limit; the first refusal is one active binding above the
limit. The clone timer stops before post-timer
byte/owner/reference validation; the cloned owner remains live through the
allocator snapshot and is dropped before receipt encoding.

The retained exploratory executables are tied to their source manifests and
preimages. An isolated source-helper commit moved `HEAD` after the attach
snapshot; that drift is recorded without treating the later commit as the
executable's source identity. Source bytes and manifests are the reproducible
authority for those bundles.

Validation performed on the gated adapter scaffold:

- standalone harness dependency resolution is locked and offline;
- the adapter is written against the current callable API;
- the release adapter compiles against an isolated candidate overlay;
- one-sample semantic smoke checks pass for all 69 acceptance and 7 exploratory lanes (not retained as performance evidence);
- raw receipts carry both FNV-1a and SHA-256 corpus identities;
- shell and Python syntax checks pass;
- the approved source pin is `ac288a303264f9ea0bb4081baa44031bee5b79a7`, and
  the committed-input guard rejects source files that are untracked or differ
  from that Git tree;
- the runner refuses an unfrozen run before creating a target;
- the runner refuses the semantic-review gate after the freeze flag alone; and
- an explicit repository `target` path is rejected by the safety guard.

Earlier scaffold reviews independently repeated Python/JSON/shell syntax checks, the unfrozen
runner refusal, a locked offline harness check in newly created temporary
Cargo targets, and Rust formatting. The current release adapter built against
retained source manifests, and the exploratory attach and shared/distinct
detach receipts passed allocator, semantic, and exact-output checks. The full
runner remains gated until the callable lifecycle API and semantic checks are
frozen; no production files were written by this profiler.

The current root review has eleven adversarial Python verifier tests. They
reject missing/nonempty stderr, missing/failed/duplicate exit-status records,
changed or malformed executable digests, incorrect process identities, and
input-size drift despite matching claimed hashes, and native receipts that
disagree with the documented producer fixture. The separate nine-test runner
suite proves missing or reused external output refusal, checkout/target/symlink
separation, symlinked-checkout resolution, and owned-target cleanup while
retaining failed-run diagnostics. It also checks that every explicitly declared
manifest extra exists in Git HEAD, catching missing inputs before a build.
The embedded CPU and memory metadata programs are executed against fixed
`/proc`-shaped inputs, so quoting errors cannot silently produce unavailable
host fields.
Namespace tests reject missing, malformed,
inconsistent, or changing first-refused binding counts. Caller-limit tests
reject missing, floating-point, boolean, nonpositive, and cross-process-drift
ceiling values. A separate five-test source-snapshot suite exercises real
temporary Git trees, including untracked, assume-unchanged, and tampered
path-source inputs. All 25 Python tests pass. Shell syntax and three early
refusal gates (unfrozen, unwired, insufficient process count) also pass without
creating a Cargo target. These checks are verifier evidence, not lifecycle
performance measurements or production acceptance.

The refusal-path review is now corrected. The Rust adapter emits a distinct
typed refusal only after classifying the public API error and confirming exact
source-byte preservation. Unexpected acceptance, source mutation, setup
failure, and unrelated API errors remain hard failures. Five adapter unit tests
cover those negative cases, exact opaque-fragment checks, registry counts, and
a validated external-edge refusal; the retained candidate build and one-sample
semantic smoke pass all 69 acceptance and 7 exploratory lanes. The native lane
also asserts the documented producer SHA-256 before emitting a receipt. These
checks are scaffold validation, not lifecycle performance measurements or
production acceptance.

The targeted coverage follow-up is now implemented in the owned scaffold, with
measurement gates still closed. The registry is 69 acceptance lanes plus 7
exploratory lanes (76 total). The isolated candidate build and one-sample
semantic smoke pass all 76 lanes, including strict attach/final-detach,
same-picture and two-worksheet attach/detach, incoming-edge retention, exact
opaque-fragment checks, derived namespace refusal, captured-owner clone, and
directly constructed shared-final fixtures.
The same-picture and multisheet lanes also require byte-stable reopen after
attach and restore, all-anchor geometry equality, and opaque-fragment checks
after attach. Composite caller-limit receipts retain each requested ceiling;
the verifier checks the complete ordered set and cross-process agreement. The
source guard binds all 15 files from `ac288a303` plus the reviewed OPC/source
prerequisites and contextual `svg_blip.rs` blob. This is readiness evidence
only; it contains no timing or performance claim.

The clean-checkout follow-up at `a556f44da` passed all 76 correctness lanes
using the newly retained native fixture without an external symlink. It also
found that the runner still listed the untracked local `docs/GOAL.md` as a
manifest extra. The manifest now binds the committed ADR 0001 and ADR 0005
alongside this profile's committed requirements. The local goal remains task
context; it is not required to build or execute the harness. The new declared
input test reproduced this missing-file failure before the runner correction.

## Sealed semantic capture at `cd5dc456b`

The [retained bundle](results/clean-cd5dc456b/README.md) contains 69 acceptance
lanes, three fresh processes each, two warmups and 20 measured samples per
process. Root and independent review verified all 207 receipts, 4,140 samples,
empty stderr files, successful GNU time exits, 4,870 manifest inputs, exact
semantic/readback outcomes, expected typed refusals, and unchanged source and
binary identities. The 635 raw files were retained byte-for-byte.

This bundle is approved for semantic and scaffold evidence only. Its original
CPU/memory provenance fields were unavailable because of the runner quoting
defect; the explicit post-run host supplement does not replace measured
provenance. Performance and optimization claims require a fresh capture under
the `e7b80f056` correction and an uncertainty summary. Raw timing data and the
generated report remain intact for audit, with that limitation stated in the
bundle README.

## Sealed baseline capture at `ab954d91d`

Root and independent review approved the [fresh capture](results/clean-ab954d91d/README.md)
for reproducible capture integrity and descriptive measurements. Its 69 lanes
contain 207 fresh processes and 4,140 measured samples, including seven typed
refusal lanes. Both source manifests agree; all 4,871 inputs and 19 pinned
production paths passed source checks. The before/after executable digest is
stable. Host CPU and memory metadata were captured during this run.

The 635 capture files and matching root verification receipt are retained
byte-for-byte. The uncertainty Markdown and JSON recompute exactly from the
raw samples and explicitly limit n=3 ranges to descriptive observations.
The current scaffold suite passes 36 Python tests. The disposable Cargo target
was cleaned; source and raw evidence remain available for review. This closes
the earlier capture's missing in-run host provenance and uncertainty evidence
for the new run only. It does not turn either capture into a comparative
optimization, native application acceptance, or scaling claim.
