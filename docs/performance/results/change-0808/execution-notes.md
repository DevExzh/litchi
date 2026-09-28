# 0808 execution notes

The previous goal turn made concrete progress: it committed the current CPU
profile and identified event handoff as a measured experiment target. Root
rechecked HEAD and working-tree state, verified the unchanged 9,196-file
production census and all 35 prior-read architecture inputs, then delegated
archive-only coding, independent semantic review, workflow-driver preparation
and independent offline analysis. Production remains unchanged during baseline
preparation. No agent runs Cargo, rustfmt, workloads or profilers.

Root resumed by checking HEAD and status: `d28e3dc702` is still current;
production remains unchanged and the three unrelated files retain their
recorded hashes. The six inherited probe files are byte-identical to final
0806. Root formatted the candidate archive with the repository configuration,
verified both buffered oracle functions are byte-identical, regenerated a
production-path patch, and refreshed the candidate manifest. `git apply
--check` passes. The first ad hoc oracle comparison used a string/bytes mix
and stopped before mutation; the corrected comparison passes. The initial
archive-relative patch was not application-ready and was replaced before any
build or application.

Pre-build driver review found an after-build target-existence contradiction,
a before/after source comparison contradiction in paired captures, a
qualification acceptance gap, an overbroad package list, and one truncated
inheritance hash. These are preparation findings, not workload failures or
candidate amendments. All must be corrected before driver freeze and the
first build; no measurement plan or adoption threshold changes.

Baseline build completed all three retained binaries. Native/profile builds
retain eleven unused-helper warnings from the inherited probe; the all-features
probe quality lane passes formatting, all 36 tests, and warning-denied Clippy.
All eighteen before qualification processes completed. Root independently
compared every source/output descriptor and the full verification maps against
seal-bound 0806 reports; all are identical, without importing historical timing.
The developing reader's qualification lane was corrected to use the allocation
binary actually used by the frozen capture driver. An earlier root schema
concern came from stale summary text and was retracted after inspecting the
actual inherited source and reports: their schema/tool remain
`litchi.pptx.public-workflow-probe-0806.v1` / `public-pptx-probe-0806`.

Candidate application followed accepted qualification. Production gates 1–3
pass: formatting, all-features/all-targets checking, and 1,241 tests (zero
failures, three ignored across 85 result groups). The new three differential
tests all pass. Gate 4, warning-denied Clippy, fails at three existing
`.err().expect()` expressions in `opened/tests.rs` (lines 464/538/557).
Root compared that whole file with `git show` at the recorded base: exact
identity. The driver stops, leaving rustdoc and boundaries unexecuted.

The explicit decision stops this batch before any after build, paired native
capture, paired allocation capture, or profile. The original frozen protocol
is not rewritten to claim completion. `restore_candidate.py` restores the
full baseline manifest exactly; all 9,196 production files and unrelated3
hashes remain unchanged. The candidate is deferred pending a separate baseline
lint repair, without a performance rejection or adoption claim. Supplemental
terminal validation and cleanup describe the actual 18 reports/18 samples and
three baseline binaries. The next experiment must use fresh paired builds.

After restoration, root ran one separate baseline Clippy control with the
same package/features/targets/warning policy. It exits 101 with the exact same
three source locations. `baseline-clippy.json` and its raw log retain this
control; it is not a successful gate, candidate retry, or performance sample.

Root independently verified 105 nested artifact descriptors, all frozen driver
inputs, the 9,196-file source chain, and unrelated-file identities before
cleanup. The supplemental early-stop cleanup then verified and removed the
three baseline binary copies and their owned target; raw evidence remains.

The post-cleanup qualification audit and independent early-stop validation
replay successfully. Root removed five unexecuted full-trial reader drafts;
the frozen workload drivers, accepted qualification reader, and actual terminal
validator remain. No unexecuted full-trial result is represented as complete.

The staged blob audit covers exactly 135 owned paths. The generic staged
whitespace check reports one extra empty EOF line in the frozen `custody.py`.
Its pre-build bytes are preserved. All other staged paths pass the normal
whitespace check; that one helper passes with only `blank-at-eof` excluded.
This archival whitespace exception changes no source, driver, policy, or result.
