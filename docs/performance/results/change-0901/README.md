# Change 0901 evidence

Scoped current-source XLSX publication Callgrind evidence for the
[report](../../0901-xlsx-publication-instruction-profile.md).

`plan.json` freezes source/binary/build/tool identities, the four-child schedule,
owner names, scope and native context. `capture.py` launches only the retained
binary. `capture/00-medium` through `03-medium` preserve each command/start,
raw report, semantics, stdout/stderr, six numbered dumps and zero-cost terminal
dump. The first child passed `admit_first.py` before later captures; its receipt
is `first-child-admission.json`. Capture directories are write-once. Do not
rerun acquisition in this packet; use a newly frozen packet for new data.

The retained 0845 `reader_v2.py` parses raw Callgrind cost/edge records, including
source-file switches. `analyze.py` is the initial helper, used for a separate
first-child offline rehearsal. `analyze_v2.py` strengthens final capture
metadata, semantic, stderr, and exclusive-cost checks. Canonical analysis is
raw-only; no ambiguous merged annotation rows are needed.

Replay from the repository root to a fresh output path:

```sh
python3 -B docs/performance/results/change-0901/analyze_v2.py --output /tmp/0901-analysis-replay.json
cmp docs/performance/results/change-0901/analysis.json /tmp/0901-analysis-replay.json
```

Remove only the duplicate replay output you create. The exact executable is
locally retained in 0894; source/locked build receipts support reconstruction
elsewhere without promising byte-identical rebuilt binaries. Raw elapsed and
resource observations from the instrumented processes are not native timing
or memory evidence. All brk-overflow warnings and repeated forked-probe stderr
collection totals are retained. Review and source audit are independent of the
root final closure receipt. Cleanup receipts identify owned temporary removal;
no worktrees or build targets were created.

An optional reviewer annotation invocation failed on the source-path display
format before producing a full analysis. Its partial annotation exports and
`annotation-attempt-failure.json` are retained; they are not authoritative
analysis. `root-replay-cli-failure.json` records a separate rejected invocation
using a v1-only flag; it created no output and ran no workload. The corrected
raw-only invocation and byte-identical replay pass. Use an absent output path:
the retained analysis helper can overwrite an existing output file.

`helper-version-history.json` records the offline helper development sequence.
The exact helper used by the failed annotation attempt is preserved separately
as `analyze_v2_annotation_attempt.py`; the final raw-only helper is hash-bound
by `replay-cleanup.json`. Original tool-output whitespace is preserved, with
explicit exceptions in `whitespace-check.json`; editable text passes the check.
`semantic-source-replay.json` freshly rechecks source identity and all four
reports against the native context without launching workloads.
