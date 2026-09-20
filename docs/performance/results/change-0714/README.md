# 0714 current DOCX publication attribution

This packet refreshes four ordinary DOCX phase distributions and separately
traces the atomic publication transaction on generated-medium and admitted
NumberedList corpora. It changes no production or harness source and makes no
before/after performance claim.

- `plan.json` and `capture-freeze.json` freeze 16 native children (two repeats,
  two corpora, four phases; 100 samples and ten warmups) and eight separate
  trace children (two repeats, two corpora, atomic/counting controls; one
  sample, no warmup). Repeat two reverses phase and corpus order; CPU 12.
- `build.json`, `source.json`, `host.json`, exact argv, source-census digests,
  fixture identities and per-child receipts bind all raw reports and traces.
- `analyze.py` reuses the strict 0709 report/decoded-manifest validators and
  0713 semantic normalizer. It recomputes native statistics, checks native/trace
  parity and lists every repeat spread over 5%.
- `trace_analysis.py` separates four setup publications from the fifth measured
  atomic publication; counting controls have no fifth publication. It follows
  exact temporary/destination/parent paths and descriptors, preserves raw line
  evidence, and excludes readback, cleanup and report-output I/O.
- `negative-checks.json` records fail-closed checks on deliberately corrupted
  in-memory inputs or temporary trace copies; retained captures are unchanged.
- `reused-verification.json` explicitly reuses 0713 correctness and six source
  evidence gates because the complete source census is identical. The 4,995
  tests, 92 passed/46 ignored doctests, and precise Clippy scope are historical
  exact-source evidence, not new runs. Seven preexisting PPTX/XLSB test Clippy
  violations remain disclosed. Current evidence/parser/doc checks are fresh.
- `cleanup.json` binds the removed executable by hash/size after deleting only
  the three owned 0714 scratch roots. `audit.py` replays the analysis after
  cleanup; `artifact-manifest.json` seals every packet file except itself.

From the repository root:

```sh
python3 -B docs/performance/results/change-0714/audit.py
python3 -B docs/performance/results/change-0714/artifact-seal.py --check
```

Capture scripts refuse overwrites. A new measurement needs a fresh packet and
scratch namespace. Trace durations are perturbed observations and cannot be
substituted for native latency. Independently sampled phase quantiles cannot
be added or subtracted to reconstruct lifecycle. The temporary-file sync,
atomic replacement and parent-directory sync remain part of the publication
contract. No hardware/device cause, cold-cache, RSS, allocation, throughput,
scaling or native Office compatibility claim follows.
