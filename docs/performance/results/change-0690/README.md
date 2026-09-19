# Change 0690 evidence packet

Lookup-only ASCII CFB key experiment after baseline `01d83efcf`.
Status: final candidate retained with explicit costs and tail limitations.
See [review.md](review.md) for the independent disposition.
See [the change record](../../0690-cfb-borrowed-directory-lookups.md).

## Reproduction and scope

Use the baseline revision plus copies of these drivers, then the candidate
revision. `build.py PHASE` builds five unchanged 0684/0686 probe crates in a
single serial lane and freezes binaries in separate before/after directories.
Adapt `/home/zhuhe/code/litchi-target-0690*`, the profile path and CPU 12
consistently on another machine. Cargo uses two jobs; incremental durations
are build logistics, not clean-build performance claims.

For each phase run `measure-native.py`, `measure-costs.py`, `measure-corpus.py`,
`measure-repeat.py`, `measure-diagnostics.py`, `profile.py`, and
`inspect-assembly.py`, each with `baseline` or `candidate`.

- Native: 12 cases × two sources × six A/A+ABBA legs × 100 fresh owners,
  eight queries each, after three warmups. 14,400 owners / 115,200 queries.
  Outcome projection and source construction are outside query timers.
- Allocation: 96 groups × three repeats × two binaries = 576 captures.
  Counted source I/O: 12 complete opening/eight-query routes. Instrumented
  timings are excluded from native latency results.
- Repeated queries: 12 groups × six legs × nine process samples, each running
  50,000 queries after two preparatory queries. All use the default 2 MiB
  limit. These are process loop means, not individual latency samples.
- Diagnostics: eight groups × two lengths × three repeats × two binaries.
  Subtract N=10 from N=100,010, then divide by 100,000. Counters include
  whole-process setup and wrappers. RSS is the native child's `time -v` value.
- Corpus: all 126 real XLS files in owned/file modes and the full generated
  70,001-cell visitor/count/digest comparison.
- Profile: two million repeated owned 54016 queries, perf 997 Hz. Assembly
  and section outputs retain exact bytes and hashes.

`audit-native.py`, `audit-costs.py`, `audit-extra.py` and `audit-final.py`
verify raw/source/binary/probe/fixture bindings, outcomes and comparisons.
`summarize.py` emits complete median and descriptive mean/tail triggers.
Bootstrap intervals do not eliminate shared-host drift. Neither a 100-owner
p99 nor a nine-process loop p99 establishes a population tail guarantee.

Run `run-integration.py` with
`LITCHI_GATE_OUTPUT=docs/performance/results/change-0690/final-verified`.
`check-consumers.py` covers DOC/PPT, `run-evidence.py` runs repository evidence
and boundary checks, and `final-doc-gates.py` checks the finalized report.
Production and test sources are frozen before final builds/checks/captures.
There is no iWork, native Office, cold-device, remote, cross-platform or
concurrency performance claim. `performance_claim: none`.

`lookup-probe/` is a new public-API legacy `OleFile::stream_len` control. Run
`build-lookup.py PHASE`, `measure-lookup.py PHASE`, then `audit-lookup.py`.
Twelve named cases cover widths 1/31/257, exact/mixed ASCII, Unicode equivalence
and fallback, missing/invalid names and nested storage. Each phase uses frozen
binaries. Nine processes per leg execute 250,000 lookups after 1,000 warmups:
648 case/process records and 162 million calls across A/A and A/B/B/A.
The ASCII queries are generally six bytes; widths measure tree size, not
key-length scaling. Fixture creation, serialization, parsing, hashing and
reporting are outside the timer; fixture lengths and SHA-256 hashes are retained. This supplements
the SharedOleFile end-to-end XLS captures. Probe development initially failed to compile because a CLI case name mixed
borrowed and owned strings; this was corrected before any baseline timing.
The final source, successful build, lockfile and smoke output are retained;
the earlier draft was not archived. A driver case-count assertion was also
corrected from ten to the actual twelve cases before retaining any timing.

`measure-followup.py` and `audit-followup.py` retain four flagged native groups
at 1,000 fresh owners per leg after ten warmups: 24,000 owners and 192,000
queries. Deterministic gzip preserves every raw value outside probe timers.
Original captures remain separate. The missing-query cost and Simple file
warm-mean p99 flag persist; the original 45365 workflow anomaly does not recur.
All phase statistics and control drift remain in the supplemental comparison.

`initial-lookup/` retains the original supplemental capture and probe with an
unused-Result warmup warning. `rebuild-lookup.py` restores the three baseline
CFB source files temporarily, builds/measures the corrected standalone probe,
then restores every candidate byte in `finally` before building/measuring the
candidate. Run it only with no concurrent source editor. `build-lookup.py`
uses `RUSTFLAGS=-D warnings`. The main XLS binaries/captures are unchanged.
`audit-lookup.py --initial` independently checks the original capture against
its archived probe; plain `audit-lookup.py` checks the final corrected capture.
Do not pool those runs. The final source/fixture identities and paired results
reproduce the same scoped disposition.
