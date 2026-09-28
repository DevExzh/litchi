# 0808 — direct PPTX event handling deferred at the quality gate

The direct-event candidate is archived, and production is restored exactly to
`d28e3dc702`. Its semantic tests pass, but warning-denied Clippy stops on three
pre-existing test expressions outside the candidate file. No paired timing,
allocation comparison, or Callgrind profile was run. This is an incomplete
performance qualification, not a measured performance rejection or speedup.

The [packet](results/change-0808/README.md) preserves the original frozen
experiment, baseline qualification, candidate archive, partial quality logs,
[decision](results/change-0808/decision.json), and exact restoration witness.
The original 310-report/6,718-sample plan remains a plan; only the eighteen
before qualification reports and eighteen measured samples were executed.

## Candidate and semantic checks

Following [0807's current CPU evidence](0807-pptx-current-capture-profile.md),
the candidate lets the notes scanner consume `Reader<&[u8]>` events directly
and delegates namespace handling to the same public `NamespaceResolver`.
Pending scopes pop before reads; Start/Empty namespace pushes precede existing
node/depth checks; Empty/End scopes pop on the next read. Namespace errors
retain their conversion through `quick_xml::Error` and the shared XML adapter.

The single production allowlist is `crates/litchi-pptx/src/notes/codec.rs`.
The buffered scanner and element-inspection oracle remain byte-identical.
Three focused tests cover nested and empty namespace scopes, reserved-prefix
errors before depth errors, and declaration-cap precedence. All pass within
the full PPTX all-features test run: **1,241 passed, zero failed, three ignored**
across 85 test-result groups. The ignored tests are not claimed as executed.

Candidate formatting and all-features/all-targets checking also pass.
Warning-denied Clippy then reports `clippy::err_expect` at `opened/tests.rs`
lines 464, 538, and 557. Root verified that this entire file is byte-identical
to HEAD. A separate post-restoration baseline Clippy control reproduces the
same three diagnostics with exit 101, confirming the independent blocker.
The [failed log](results/change-0808/quality-after/03.log) is retained;
the driver stops there, so rustdoc and the crate-boundary gate were not run.
There is no claim that the six-gate production quality suite passed.

## Baseline evidence and custody

Three baseline binaries were built from the 9,196-file source census. Native
and profile builds each retain eleven inherited unused-helper warnings; the
allocation build has none. The separate all-features probe quality lane passes
formatting, all 36 probe tests, and warning-denied Clippy.

All six probe files are byte-identical to the final repaired 0806 probe. The
schema/tool remain `litchi.pptx.public-workflow-probe-0806.v1` and
`public-pptx-probe-0806`; the profile wrapper symbol remains
`namespace_uri_probe::capture_region_0793` but was not profiled in this batch.

The eighteen before-only qualification runs cover six shapes crossed with
capture, commit, and lifecycle. Their source, output, text, fixture, and
extension-preservation oracles match the sealed 0806 qualification evidence.
Root additionally compared the complete verification maps. These comparisons
import no historical timing and supply no before/after performance estimate.
Qualification was accepted before the candidate was applied.

All 35 architecture inputs were checked against their recorded hashes. After
the quality failure, the one-file candidate was restored and all 9,196 source
hashes match the baseline. The three unrelated working-tree files are preserved.
The owned target and three baseline executable copies are removed only after
identity verification; the terminal validator replays the retained qualification
and partial quality evidence after cleanup and checks staged/committed blobs.

## Next work and limits

Repair the three independent baseline Clippy blockers first. Then start a
fresh experiment from that repaired source, with new baseline and candidate
binaries and the same semantic obligations. This packet does not authorize
bypassing the quality gate or reusing its baseline timings as a paired control.

No public API, dependency, production optimization, supported-format capability,
or performance claim is retained. No native Office, cross-platform, cold/range,
concurrency, cross-format, or full CRUD coverage is added. OLE2/OOXML work and
the broader performance goal remain incomplete; iWork is excluded.
