# Log sections for change 0626

Four paragraphs for the coordinator to insert at the top of the program logs,
each in the style of that file's newest section. This change does not edit
`HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` or `ADR_COMPLIANCE.md` itself. The
relative links below are written from `docs/performance/`, their destination,
not from this file.

---

## For `docs/performance/HOTSPOTS.md`

## 0626 — the CI smoke check gets a real baseline, and the comparison it was already running turns out to have been failing closed for 273 commits

Evidence gap 1 of the 0587 survey, ranked highest of twelve, was that the
`smoke` job of `.github/workflows/perf-baseline.yml` "builds `baseline` as a
byte-copy of `current` with only the revision label changed, then asserts the
comparator passes against itself" — a working plumbing check that cannot detect
anything, because both sides are the same measurement. The `full` job now also
captures the bounded allocator report the smoke job captures (the same release
`litchi-perf-baseline-alloc`, `--warmup 3 --samples 15 --case
opc_file_eager_open --filesystem-cache warm,cold-requested`) and ships it, with
a descriptor naming the run, in the existing baseline artifact; the smoke job
lists recent successful runs on `main`, downloads the newest such artifact and
compares against it, falling back to the old self-comparison — **labelled, in the
job summary and the annotation, as a plumbing check that detects no
regression** — only when the fetched report is not bound to the same comparator
policy identity, case/corpus key manifest digest, harness binary profile and
runner labels, at a distinct revision, from a clean worktree. The survey's own
premise did not survive the work: that self-comparison **has not run since
`126c4a8b2`**, 273 commits back, because the allocator harness began emitting
`tool.allocator_counter_revision` while the checked
`perf-regression-policy-allocator-v1.json`, last edited two weeks earlier,
omits it — and `perf_compare` compares the *complete* tool identity, so every
freshly captured allocator report is rejected with `baseline.tool does not match
the policy tool identity` and the step exits 2. Reproduced on two real reports
from the release harness and fixed by pinning the field the harness actually
emits, with a test that asserts the pinned string against the harness source so
the two cannot drift apart silently again. Five end-to-end scenarios over those
two real reports (distinct revisions `7082a1a3f` and `77220b62b`, both clean,
both pinned to CPU 21) behave as designed: no artifact → self-comparison,
plumbing pass; compatible reference → real comparison, pass; **allocation
counters raised 10% → comparator regression, reported, job exit 0**; reference
CPU model changed → comparator invalid, classified as infrastructure drift, job
exit 0; reference over another corpus → falls back naming both digests. The
comparison compares twenty deterministic allocation counters and withholds every
latency claim, and `enforcement: advisory` keeps it off the merge gate per
`docs/GOAL.md` deliverable 8; only a broken plumbing check fails the job.
`performance_claim: none`; no file under `crates/` or `tools/perf-baseline/`
changed. OLE2/OOXML remain active; ODF is deferred until completion and iWork
excluded. [Record and limitations](0626-perf-ci-smoke-baseline-fetch.md);
[retained evidence](results/change-0626/README.md).

---

## For `docs/performance/GOAL_AUDIT.md`

## 0626 — deliverable 8's smoke check starts comparing against history, and stays off the merge gate on purpose

Record: [0626](0626-perf-ci-smoke-baseline-fetch.md).

`docs/GOAL.md` deliverable 8 asks for "a lightweight, stable performance smoke
check suitable for CI, plus a fuller manually triggered or scheduled benchmark
workflow", and adds "do not make noisy cloud-hosted microbenchmarks a hard merge
gate until variance is understood". Both halves existed; the first half was
comparing today's report with itself. It now compares against the last
successful `full` run's artifact when one is bound to the same policy identity,
corpus key manifest digest, harness binary profile and runner labels, which is
the first time any push or pull request in this program has been measured
against prior history at all. The second half is honoured structurally rather
than by assertion: the comparison runs under the allocator policy, so
`perf_compare` reports `latency_claims: withheld_instrumentation` and compares
zero latency results, and the twenty metrics it does compare are deterministic
allocation counters; on top of that `enforcement: advisory` means a regression
or an unusable comparison annotates the run and does not fail it. The audit rows
this leaves open are named in the record: nobody has yet observed a GitHub
Actions run of this workflow, because neither `gh` nor `actionlint` is installed
on the program host, so the fetch step's shell has never executed; whether a
hosted runner's build identity is stable enough between runs for the fetched
mode to fire in practice is exactly the variance deliverable 8 says must be
understood, and it is unknown; and the reference is one case over one pinned
corpus, not the 201-result release matrix a pull-request job cannot afford.
Deliverable 6's machine-readable results gain a small, honest addition — the
selection, comparison and classification of every smoke run upload as
`container-performance-smoke-comparison-<run id>`, each stating in its own text
whether it detected regressions or merely checked plumbing. Two audit
corrections are owed to earlier entries: change 0587's evidence-gap table
described the smoke self-comparison as "a real, working plumbing check", which
was true when written but had not been true since `126c4a8b2`; and change 0421's
"policies opt into the new identity once both sides are freshly captured"
applies to the CI allocator policy, where both sides always are, and was never
acted on. No production code, contract, limit or refusal is touched. OLE2/OOXML
remain active; ODF is deferred until completion and iWork excluded.

---

## For `docs/performance/REPORT.md`

## 0626 — what the smoke job now compares, and what it had been failing to compare

Change 0626 closes evidence gap 1 of change 0587 and, in doing so, repairs a
silent CI failure the survey had classified as working. The `full` job gains one
step — the same bounded allocator capture the smoke job already runs, written to
`target/perf/allocator-baseline.json` with an
`allocator-baseline-descriptor.json` beside it, both added to the existing
`container-performance-baseline-<run id>` artifact — and the smoke job gains
four: a best-effort `gh run list` / `gh run download` of the newest successful
run on `main` that published one, a selection step, a comparison step that
captures the comparator's exit status instead of failing on it, and a reporting
step that classifies the verdict and writes the job summary. All of the decision
logic lives in `tools/perf_smoke_baseline.py` (five subcommands) rather than in
workflow heredocs, and `tools/test_perf_smoke_baseline.py` covers every branch
with 121 tests across thirteen classes, nine of which assert the workflow's own
wiring textually. Verification used two **real** release
`litchi-perf-baseline-alloc` reports, each captured on a clean worktree pinned to
CPU 21 at a distinct revision — `7082a1a3f` standing in for a prior full run and
`77220b62b` for the commit under test — driven through the whole pipeline by a
retained script, once per scenario: no artifact fetched → `self_comparison`,
comparator pass, `plumbing_pass`, job exit 0; a compatible reference →
`fetched_reference`, comparator pass, `reference_pass`, job exit 0; the current
report's allocation counters raised 10% → comparator **regression**, exit 1,
`reference_regression`, **job exit 0** because enforcement is advisory; the
reference runner's CPU model changed → comparator invalid, exit 2,
`reference_environment_drift`, job exit 0; a reference captured over another
corpus → rejected before comparison, naming `36f44718…` against the policy's
`debb7009…`, and falling back. Over the two real legs the comparator reports
pass, 2 matched results, 20 compared metrics, 0 regressions,
`latency_claims: withheld_instrumentation`, 0 latency results compared and 2
excluded. Against the base commit's own copy of the allocator policy the same
pair is rejected outright — `ComparisonInputError: baseline.tool does not match
the policy tool identity` — which is what the CI step has been doing since
`126c4a8b2` added `tool.allocator_counter_revision` to the harness while the
policy stayed markerless; the fix pins `serialized_region_peak_v3`, the value the
harness source emits, and a test ties the two together. Gates: `cargo fmt --all
--check` clean; the workflow parses under PyYAML 6.0.3; 466 tests across the nine
modules the smoke job runs, OK; `tools.test_perf_baseline_source_policy` 14 OK;
`check_perf_claims.py --mode strict` 10 claims;
`check_report_claim_classification.py` 167 rows;
`validate_crud_coverage_index.py` 15 categories and 33 selectors;
`non_iwork_gate.py verify` 45 bulk tree roots and 35 facade safe trees. One
pre-existing failure, reproduced unchanged at the base commit:
`tools.test_perf_claims`'s registry structural test names two ABBA evidence ids
its expected set omits. `actionlint` is not installed on this host and the
workflow was not linted by it. `performance_claim: none`; no file under
`crates/` or `tools/perf-baseline/` changed.
[Change and limitations](0626-perf-ci-smoke-baseline-fetch.md);
[retained evidence](results/change-0626/README.md).

---

## For `docs/performance/ADR_COMPLIANCE.md`

## 0626 — no boundary is engaged; what the checked policy identity now pins, and why that is the strict direction

Record: [0626](0626-perf-ci-smoke-baseline-fetch.md).

No ADR is engaged and no accepted hash changes: the change is confined to a CI
workflow, two `tools/*.py` files, one new policy document, one pinned field in
an existing policy document and one documentation file, and `git diff 7082a1a3f
-- crates/ tools/perf-baseline/` is empty. No contract, limit, refusal, output
byte, public type or `unsafe` boundary moves, and `tools/perf_compare.py` — the
fail-closed comparator itself — is not modified; what changed is what the smoke
job does with its verdict. The one compliance-shaped decision is the pin of
`tool_identity.allocator_counter_revision` to `serialized_region_peak_v3` in
`perf-regression-policy-allocator-v1.json`. It narrows rather than widens: a
report captured by an allocator binary older than `126c4a8b2` can no longer be
compared under that policy, which is what change 0421 prescribed —
"legacy markerless policies can still replay historical pairs… both sides must be
freshly captured with corrected instrumentation before a policy opts into the
new identity" — for the one policy where both sides always *are* freshly
captured. `perf-regression-policy-xlsx-allocator-v1.json` and
`perf-regression-policy-opc-source-materialize-allocator-v1.json` are
deliberately left markerless, because retained evidence reads them and changing
them would reinterpret that evidence in retrospect. Three further properties are
worth recording because a CI wiring change is easy to assume harmless: the
fetched artifact is **not trusted on its own word** — the case/corpus key
manifest digest is recomputed from the fetched report and must match both the
policy's `expected_result_keys_sha256` and the descriptor's copy, and the
descriptor's run id must equal the run the artifact was downloaded from, so a
stale or mixed download cannot be compared; the fetch step cannot fail the job,
carrying `continue-on-error: true`, running under `set +e` and recording its
outcome in a typed status document rather than an exit code, so a GitHub API
outage degrades to the labelled plumbing check instead of a red build; and the
advisory enforcement has exactly one exception, a broken self-comparison, which
fails the job under either setting because a comparator that cannot compare a
report with itself is a tooling defect and not a measurement. The manual
`reference-regression` job's fail-closed behaviour on a regression or identity
defect is unchanged. `performance_claim: none`. OLE2/OOXML remain active; ODF is
deferred until completion and iWork excluded.
