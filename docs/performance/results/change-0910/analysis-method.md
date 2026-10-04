# 0910 analysis method

`analyze.py` is an offline, fail-closed reader for the 0910 generated PPTX
capture. It reads `capture-plan.json`, the selected build and quality receipts,
and all 24 `capture/block-*/command.json` receipts. Every receipt descriptor is
resolved inside this packet and checked against the current bytes and SHA-256.
The receipt's route, block, position, binary, source manifest, plan, logs,
artifacts, report, and `/usr/bin/time` process metrics must agree with the
plan. Reports must contain 30 warmups, 20 measured samples, the expected 229
part / 445 member corpus, the frozen source digest, 64 replacements, and all
semantic, stable-output, sink, and source-immutability gates. Each raw sample
vector is checked again, including the bounded 64 KiB sink counters,
`total_ns >=` the five phase durations, output identity, and ordered indices.

The plan supplies twelve serial AB/BA pairs. The analyzer pools all 240
measured samples per route for descriptive p50, p95, and p99 statistics and
reports medians for every timed phase and total. P50 is the
midpoint of the two central ordered observations; p95 and p99 use nearest
rank. For each block it computes Python's paired-process median for the
materialized samples divided by the corresponding owning median. The reported
95% interval is a deterministic paired bootstrap of the twelve block ratios:
10,000 resamples of twelve complete blocks, Python `random.Random(910)`, and a
nearest-rank percentile interval over bootstrap medians. Order drift compares
the six AB blocks with the six BA blocks per route. Maximum RSS is retained per
process and summarized descriptively; its scope includes corpus generation,
warmups, semantic guards, and measured samples, so it is not operation peak
live memory.

The script creates its requested JSON output only when all checks pass. It
does not invoke Cargo, the benchmark binary, replay, profilers, or external
commands.
