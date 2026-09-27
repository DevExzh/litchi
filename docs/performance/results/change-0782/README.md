# Change 0782 evidence packet

The optional borrowed PPT text slice candidate is rejected. Source is restored
exactly to the recorded base; there is no retained production change. Many
write/lifecycle and Unicode write/lifecycle violate the frozen persistent
p50 regression guard despite large ASCII payload improvements.

The packet retains 120 native, 40 allocation, ten qualification, twelve perf
and two Heaptrack processes, plus the standalone probe and six quality gates.
All 170 primary reports and 3,730 samples pass source/output, semantic and raw
text checks. Quality records 1,287 passed tests, 11 ignored, 34 suites; probe
tests pass 31 default-feature and ten all-feature tests. All 42 native spread
and 24 paired-series flags remain; allocator flags are zero.

The corrected probe is intentionally byte-identical to 0781 and retains that
tool identifier. `probe-inheritance.json` binds its five source/template/lock
files; `baseline-fixture-parity.json` binds all ten qualification identities.
Replay checks the actual historical files and their seal, not just the new
receipt. The historical timings are not part of this paired comparison.

From the repository root:

```bash
python3 -B docs/performance/results/change-0782/validate.py
python3 -B docs/performance/results/change-0782/tables.py --check
python3 -B docs/performance/results/change-0782/decision_audit.py
```

The final validator requires the packet seal and replays source, probe, build,
binary, exact command, allocation, timing, oracle, quality, observer, decision
and restoration contracts. Missing executables require exact cleanup witnesses.
The five tables must reproduce byte-for-byte. `analysis-before-disposition.json`
retains the first complete arithmetic before the final rejection record.

See the [report](../../0782-ppt-borrowed-slice-rejected.md) for complete metrics,
limitations, reproduction order and integration/cleanup closure. Reproduction
requires fresh output directories and receipts. The final disposition is in
`disposition.json`; `candidate/before` and `candidate/applied` retain both source
states. Archive notes/receipt describe the initial draft, while
`candidate/application-notes.md` records the coordinator's formatted selection.

Cold-cache, device, concurrency, physical-copy and complete CRUD claims remain
out of scope. OLE2/OOXML remain active, ODF deferred, and iWork excluded. No
coverage registry status is promoted.
