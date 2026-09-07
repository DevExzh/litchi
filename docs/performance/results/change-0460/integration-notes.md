# Integration and validation scope

The baseline is revision `c16fd8cb5`; exact compiled source manifests and retained
binary identities are authoritative. Candidate files are supplied in
`source-delta.json` as complete `.txt` artifacts for both source epochs. The
retained patch describes the measured experiment even if the final decision
rejects it. Evidence text copies introduce no additional Rust build inputs.

The private staging path feeds settings, declaration, page-metadata and
source-fragment state machines from one namespace-aware event stream. Settings
and shared XML errors stop immediately; declaration and page errors remain
deferred in historical order. Source errors are carried inside the staged
result and applied only after metadata, MIME and styles setup. Byte spans use
the original BOM-stripped input; the BOM is emitted verbatim. Source ownership,
all candidate reopen/readback checks, patch identity, and no-op semantics remain
as before. No clocks, pools, unsafe code, dependencies or public APIs are added.

The first owner-test attempt failed compilation on two unqualified associated
function calls and three unnecessary test qualifications. Its original source
manifest, receipt and log are retained; the corrected retry is separate. The
first compiled retry passed 184 library tests and failed one new priority
fixture: its settings element was outside the structure recognized by the
independent reference. The fixture was corrected without changing production
parsing; later retries retain distinct receipts.
Correctness tests compare independent pre-refactor scanners over malformed and
namespace corpora, limits and native fixture XML, with explicit BOM byte checks.
The native fixtures are library-level checks, not Office GUI roundtrips.

Only the ODP owner changes. The candidate passes 371 release all-feature owner tests, 387 standalone
harness tests (one ignored), warning-denied Clippy across all owner targets,
warning-denied rustdoc and scoped formatting. Crate-boundary results are
recorded in the passing final receipt. ODP has no standalone fuzz target; unchanged
ODF detection/ODT parsing targets do not exercise this private staging path.
Existing differential malformed-input corpora exercise its relevant parser
boundaries. No unrelated fuzz coverage is claimed. The known unrelated Keynote
format difference from 0459 remains outside the requested non-iWork scope.

Capture and supplementary drivers were frozen before candidate binding. The
primary 24-report/720-sample A1/B1/B2/A2 lifecycle matrix drives the decision;
120 separately phase-timed large operations and two 100-sample whole-process
counter runs are diagnostic. Source/open, transaction, append, commit and sink
write remain inside the lifecycle boundary. Semantic/hash/oracle setup and
reporting remain outside that clock. Allocator and normal results are separate.
GNU-time RSS is a process-lifetime high-water mark, not regional live memory.

All captures run one workload at a time on CPU 2. This is one host, warm process
execution, owned input and sequential sink output; it establishes no cold-cache,
range-source, multiworker scaling or native application improvement. Registry
counts remain 439 selectors / 36 defaults. The full non-iWork goal stays open.
