# XLSX profile scaffold readiness

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
targets, inventory counts, namespace pressure, and refusal lanes. It must
still satisfy every semantic and refusal gate in `requirements.md` before
setting `XLSX_SVG_PROFILE_API_WIRED=1`. The three dedicated multi-picture
attach lanes and six exploratory detach lanes stage 16, 64, and 256 edits on
one drawing and validate one composed commit/reopen per count. They are
bounded cost lanes for the current per-intent rescan behavior, without a
scaling claim. The 60-lane acceptance matrix adds exact in-memory inverse and
replay-refusal checks for every anchor/size pair and a composite mixed-cap
atomic-refusal lane. The adapter repeats OPC reachability, SVG leaf, content
type, fallback, and selected-owner checks after each changed publication;
source-copy and cumulative staged-byte fields stay null until the production
owner exposes those boundaries.

The retained exploratory executables are tied to their source manifests and
preimages. An isolated source-helper commit moved `HEAD` after the attach
snapshot; that drift is recorded without treating the later commit as the
executable's source identity. Source bytes and manifests are the reproducible
authority for those bundles.

Validation performed on the gated adapter scaffold:

- standalone harness dependency resolution is locked and offline;
- the adapter is written against the current callable API;
- the release adapter compiles against an isolated candidate overlay;
- one-sample semantic smoke checks pass for all 60 acceptance and 7 exploratory lanes (not retained as performance evidence);
- raw receipts carry both FNV-1a and SHA-256 corpus identities;
- shell and Python syntax checks pass;
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

The current root review adds seven adversarial Python verifier tests. They
reject missing/nonempty stderr, missing/failed/duplicate exit-status records,
changed or malformed executable digests, incorrect process identities, and
input-size drift despite matching claimed hashes, and native receipts that
disagree with the documented producer fixture. A copied-runner execution test
also proves that cleanup preserves historical exploratory and unrelated
evidence, and removes its owned target on build failure. All seven pass. Shell syntax
and three early refusal gates (unfrozen, unwired, insufficient process count)
also pass without creating a Cargo target. These checks are verifier evidence,
not lifecycle performance measurements or production acceptance.

The refusal-path review is now corrected. The Rust adapter emits a distinct
typed refusal only after classifying the public API error and confirming exact
source-byte preservation. Unexpected acceptance, source mutation, setup
failure, and unrelated API errors remain hard failures. Two adapter unit tests
cover those negative cases and a validated refusal; the retained candidate
build and one-sample semantic smoke pass all 60 acceptance and 7 exploratory
lanes. The native lane also asserts the documented producer SHA-256 before
emitting a receipt. These checks are scaffold validation, not lifecycle
performance measurements or production acceptance.

Independent review approves retaining the scaffold with measurement gates
closed. Before final profiling, add explicit Strict-host lifecycle and incoming
edge lanes, strengthen opaque-descendant preservation assertions, derive the
namespace refusal boundary from the admitted policy, separate true captured
owner cloning from open/save work, and build shared-final fixtures without
using production detach as fixture preparation. The current 67-lane smoke
does not close these acceptance-coverage gaps.
