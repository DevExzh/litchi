# Opened DOCX edit/save provider baseline (0494)

This batch extends the existing one-paragraph edit and sequential publication
benchmark across explicit source providers. It establishes provider-specific
baselines; it does not claim a production optimization or compare those costs
to the read-only lifecycle in 0493.

The operation reuses the existing deterministic media-rich DOCX corpus and
`publish_docx_source_edit` path. Source construction and output-sink reservation
are outside the operation clock. Mandatory open, edit, commit, publication,
and document/commit teardown belong inside it. Retained output is checked after
timing for one changed paragraph and preservation of all unselected parts,
relationships, and eight media members. Exact timer boundaries and supported
provider cells are defined by the final harness and capture protocol.

The managed ordinary edit and genuine borrowed source gaps remain explicit.
This baseline must not remove the existing managed typed refusal or label a
copied byte owner as borrowed. Filesystem cold eligibility requires the existing
residency and process-I/O proof; warming an input after its final probe makes
that sample ineligible.

The source manifest includes the preexisting Keynote compilation input solely
for reproducibility. That unrelated working file, `docs/GOAL.md`, and
`docs/report/spec-gap-audit.md` are excluded from this batch's commit.
The protected `~/code/litchi-spec-gaps` worktree is untouched.

## ADR constraints

`adr-refresh.json` verifies that the complete previously read ADR set is
unchanged. ADRs 0001/0004 keep provider controls within the expert benchmark;
0002/0010/0011/0024 keep archive ownership below the format; 0003 requires
source-bound reversible edits and atomic publication; 0005 defines explicit
providers, budgets and measurement scope; 0006 requires preserve-or-refuse;
0008 and the CRUD checklist require separate opened-edit and producer evidence.
No architecture decision is changed by this harness work.

## Validation and reproduction

Build, source, command, environment, raw observation, and helper custody records
are retained beside this file. Failed development attempts remain visible;
only the final source-stable Rust gates and verified captures support the
benchmark results. Python postprocessing authenticates those retained inputs;
its own helper custody does not certify unrelated Rust edits made afterward.
The [change record](../../changes/0494-docx-edit-provider-baseline.md) links
accepted results and limitations. Accepted evidence comprises 720 warm and 120
cold formal samples, plus 36 warm and six cold pilot samples. Run
`python3 -B docs/performance/results/change-0494/verify_bundle.py verify` from the
repository root to verify the completed bundle and retained executable custody.

The full performance goal remains active. Atomic filesystem save, native
producer round trips, genuine borrowing, managed ordinary edits, concurrent
scaling, and the broader CRUD matrix require further work.

The CPU lock keeps the historical `litchi-goal-0484/cpu.lock` path deliberately:
it serializes this performance lane with its predecessors. It is advisory and
does not isolate the shared host from unrelated builds. The earlier development
cache was missing at resume, so only newly retained final executables may run
the accepted pilot and formal protocols.

For the cold selector, the timed boundary also includes `FileSource` open and
source-version fences. The warm file provider constructs that owner outside the
clock and uses the unpadded source. The cold source has a page-aligned ZIP tail.
Consequently these cells are separate baselines, not a controlled measurement
of the cost of filesystem cache residency alone. Child `/proc` RSS and VmHWM
remain distinct from GNU time's parent/waited-child process-tree maximum.

Environment receipts bind the selected build/capture variables listed in
`support.ENV_KEYS` and record tool versions and paths. They do not claim a
hermetic environment: other inherited variables, host scheduling, and the
shared filesystem remain external conditions. Raw timing and repeat variance
must be interpreted within that limitation.

Empty reads are counted by the outer `CountingReadAt` and delegated to its
immediate wrapped provider. For the range adapters, that provider returns zero
without calling its underlying physical source. Thus logical empty calls need
not equal physical empty calls or transport-service calls; nonempty range
conservation is checked separately.

## Profiler interpretation

The original `profiling-r1/profile-summary.json` is retained as a historical
attempt with a failed stack export and incorrect perf/strace parsed fields.
Use the [corrected observer receipt](profiling-counter-recovery-r5/observer-correction.json)
selected by the final bundle verifier for counter totals and stack classification. `profiling-recovery-r3` exports the
same original recording; subsequent counter corrections only reparse retained
text. No new DOCX workload samples are implied by recovery attempts.
Whole-child observer counts include setup, output oracles, and report writing.
Bare inline symbols and incomplete callchains limit stack attribution.
