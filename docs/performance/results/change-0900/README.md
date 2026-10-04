# Change 0900 evidence

Current-source XLSX one-percent edit/save phase baseline. See the
[report](../../0900-xlsx-edit-save-phase-refresh.md).

`plan.json` freezes the retained source/binary/build identities and native
schedule. `run.py` is the original helper; its first successful qualification
workload was rejected only by an incorrect optional Git-metadata assertion.
`metadata-addendum.json` and `run_v2.py` preserve and correct that failure.
No workload was repeated. `qualification/` retains two children / eighty
samples; `native/` retains twelve children / 1,200 samples. They remain separate.
Every child keeps command/start receipts, stdout/stderr, raw report and validation.

Replay from the repository root to a fresh output path:

```sh
python3 -B docs/performance/results/change-0900/run_v2.py replay --output /tmp/0900-analysis-replay.json
cmp docs/performance/results/change-0900/analysis.json /tmp/0900-analysis-replay.json
```

Remove only the duplicate replay output you created. Capture paths are
write-once and already populated: do not rerun capture commands in place.
A new acquisition needs a separate packet and freshly bound plan/source/build.
The exact release executable remains locally retained in the 0894 packet;
rebuilding from the identified source/lockfiles is required elsewhere and does
not promise identical executable bytes. Optional runtime Git/compiler fields
are unavailable because the original isolated build root was cleaned and the
fixed child PATH lacks rustc; external source/build receipts are authoritative.

`analysis.json` contains all phase totals, quantiles, shares, conservation,
corpus/resource identities and child variation flags. The independent review
and source/methodology audits are separate from the root final closure audit.
`scope-clarification.json` preserves the returned-snapshot drop boundary,
qualification recovery, and actual review timing. No historical time comparison,
allocation capture, profiler capture, build, or production change is included.

`review.json` and `independent_review.py` retain the initial independent review,
including its now-superseded prose bindings and combined metadata explanation.
Use `review-final.json` for current bindings. The independent reviewer supplied
`independent_review_v2.py`; root added exclusive output creation before its
first execution, then ran it against the final audit notes. Numerical review
logic is unchanged. Git metadata is unavailable because of the removed build
root; compiler metadata is unavailable because the fixed PATH omits rustc.
`final-review.json` is the separate root closure receipt. The first review is
historical and is not claimed to bind the subsequently corrected prose.
