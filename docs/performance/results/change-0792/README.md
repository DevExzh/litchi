# 0792 — exact empty attribute tails

Current 0791 profiles identify checked attribute iteration beneath the PPTX
notes scanner. This experiment adds an early return only for an exactly empty
raw attribute tail, after namespace, local-name, and root validation. All
attribute-bearing and whitespace-tail paths keep checked iteration. The
independent buffered scanner oracle remains unchanged.

The unchanged 0785 probe covers capture, staged commit, and full lifecycle for
tiny, medium, large, ASCII vendor, and Unicode vendor fixtures. Source/output
and semantic identities are checked against the sealed fixtures; historical
timings are not pooled. Native work uses six alternating blocks, thirty samples
and three warmups. Allocation work uses two blocks, three samples, no warmup.
Baseline qualification is separate. All commands run serially on CPU 12.

The pre-build policy requires at least one capture or lifecycle paired p50
improvement of 3% with bootstrap upper bound below one. A case regressing over
5% with lower bound above one rejects the candidate. Net live bytes, peak above
entry, allocation calls, and allocated bytes may not increase. RSS remains a
whole-child accounting observation subject to 0789/0790 limitations. Guest
profiling is deferred until the native gate establishes a useful candidate;
this batch makes no instruction-count or causal phase-share claim.

Build/capture drivers record exact commands and hashes. They refuse existing
outputs; reproduction requires a fresh packet and owned target, matching source
and lock files. The execution order is `build.py before`, `probe_tests.py`,
`capture.py qualification`, applying `candidate/applied.patch`, `quality.py`,
`build.py after`, `capture.py native`, then `capture.py allocation`.
Run only the offline replay command against this retained packet:

```sh
python3 -B docs/performance/results/change-0792/validate.py --require-final-seal --check-workspace
python3 -B docs/performance/results/change-0792/root_audit.py --check
python3 -B docs/performance/results/change-0792/quality_summary.py --check
python3 -B docs/performance/results/change-0792/sample_pair_audit.py
```

No cold/range/concurrent/native-producer coverage or broader CRUD completion
follows. OLE2/OOXML remain active; ODF is deferred and iWork excluded.

The initial formatting failure and quality-attempt-0 pointer-test failure are
retained. `test-correction.json` documents the test-only allowlist amendment,
installed Rust reference, and baseline diagnostic; `final-application.json`
binds the exact combined candidate, supplemental tests, and test correction.
The final quality attempt is authoritative; none of these corrections replaces
a primary paired measurement.

The candidate is retained: large capture improves 5.211% and lifecycle 6.937%
by median paired p50. There are 21 native process-spread flags and six metric
families with an individual paired regression over 5%; all remain in analysis.
No p50 or allocation guard fails. See the [report](../../0792-pptx-empty-attribute-tail.md)
for scope and uncertainty. Omit `--check-workspace` if unrelated workspace
files/worktrees have subsequently changed; source/document identity checks
require replay at this packet’s commit.
