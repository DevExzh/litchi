# 0897 restored fallback allocation analysis

This packet analyzes the retained 0895 baseline children 00 and 03. It launches
no workload, rebuild, or profiler. `source-identity.json` binds unchanged
production and harness source; `prior-capture-verification.json` records a
fresh offline replay of the original capture verifier.

Run from the repository with a new output path:

```sh
python3 -B docs/performance/results/change-0897/reduce.py --output /tmp/0897-replay.json
```

The reducer reads the retained receipts/stacks and performs read-only Git
identity checks. It accepts a descendant of the observed evaluation commit
only while relevant source bytes match both pinned source revisions. Output
includes the currently observed HEAD, so replay across documentation commits
can differ in that identity field while allocation results remain unchanged.

`analysis.json` is the canonical frozen result. `replay-verification.json`
records the coordinator's byte-identical replay at the same HEAD. The cleanup
receipt maps any removed duplicate replay file to canonical bytes. Original
stack line numbers, hashes and costs retain duplicate exported signatures;
identical stack text does not mean duplicate allocation events may be dropped.

`source-audit.md` proposes emitted QName output spans for a later isolated
candidate. A stack signature cannot prove one exact source line or a removable
allocation count. No latency, allocation-byte, RSS, or adoption claim is made.
