# 0531 rejected experiment: OOXML MCE namespace search

The production change is reverted. The original candidate passed the native pilot, but the final rebuild produced a different binary. A fresh complete 44-child ABBA campaign failed the frozen dense-sparse repeat 2 total p50 and mean gates (0.8937% and 0.6056% reductions versus the required 1%). The earlier passing measurements do not authorize retention.

`final-native-comparison.json` and `final-native-review.json` preserve the rejection evidence. `final/` contains the measured rebuilt candidate, while `restored/` binds the checked-in baseline production plus nine corrected regression tests and a test-only Clippy correction. `restored-quality-summary.json` records all eight passing quality gates and 4,757 successful test executions on that restored source. Historical candidate, final-candidate, and failed test/lint attempts remain intact.

Four additional final-binary profiles are diagnostic only. Original allocation, eager, reopen-confirmation, and hardware evidence applies to the initial candidate; it is not presented as validation of a retained optimization. No further candidate measurements are needed to reject the change.

OLE2 and OOXML remain the active optimization priority. ODF is deferred and iWork is excluded. See `next-ole2-review.md` for the next measurement proposal.

---

The following original campaign description documents the initial pilot and its scope; the rejection above supersedes any provisional retention interpretation.

# 0531 shared OOXML MCE namespace-search pilot

This campaign tests the one-line namespace-search candidate identified by
0530 planning attribution. The baseline is byte-identical to 0530 and 0529
final source. The candidate replaces only the exact namespace-presence byte
predicate with `memchr::memmem::find`, using the existing dependency, and adds
nine public regression tests. All input/output limit ordering and the full
MCE parser remain unchanged.

The frozen plan requires every primary XLSX shape/repeat to improve measured
workflow total p50 and mean by at least 1% and planning p50 by at least 2%.
These thresholds target a small shared-substrate change whose enclosing MCE
edge was 5.179072% of planning instructions in0530; they are not an estimate
of achievable wall time. Conditional planning profiles, allocation diagnostics,
quality gates and eager guards are required before retention. Every adverse
metric and same-build drift over5% requires individual review.

Native order is baseline1, candidate1, candidate2, baseline2, with retained
baseline binaries hash-bound under the candidate checkout for the final run.
Both primary shapes use200 samples after20 warmups. Nine guard children per
repeat use30 samples after10 warmups: seven XLSX children plus fixed media-rich
DOCX and PPTX source-backed edit/save consumers. The XLSX shape argument does
not alter the latter two fixed corpora. Native times cover measured workflows;
RSS covers the complete child. CPU2 affinity does not imply an idle host.
External Cargo activity observed in the shared workspace is retained separately.

All Rust builds, quality commands and campaign children run serially under
root coordination. Reviewers prepare source/tests and independent analyzers.
The source manifest includes the new test even before Git staging. The draft
is retained separately from its rustfmt-formatted installed source; both are
hash-bound. Reproduction must use a fresh result/target directory and exact
locked source, never write into this retained evidence bundle.

OLE2/OOXML performance remains the active goal. ODF is deferred until that
goal completes, and iWork is excluded. This pilot alone does not establish
cold-cache, range-source, native-producer, scaling or program-level completion.

The pre-cleanup verifier passed all eight components. `cleanup.json` records removal of both owned temporary paths and the bundle Python cache after checking accessible process references. Reproduce the complete sealed evidence check with `python3 -B docs/performance/results/change-0531/verify.py --strict` from the repository root.

Final status: all ten verifier components pass after cleanup, including the rejected-decision and exact report replays. `verification.json` retains the compact result. No production speedup is adopted.
