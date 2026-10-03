# 0892 current-source producer XLSX profile

This packet profiles the retained 0891 native executable after the adopted
x14ac change, with the current repository source identity frozen by
`source-identity.json`. The source audit compares the current selected library
and harness files with the 0885 archive plus the two sealed 0891 overlays. It
allows only the four separate `xml_attribute_bounds` diagnostic files from
0889; the diagnostic is not linked into the profiled binary. The retained
binary, 0891 build receipt, archive, manifest, candidate patch, quality
receipt, and adoption decision are bound by the root plan.

The native plan uses the 0886 protocol: three sequential fresh processes on
CPU 12, `cycles:u`, perf frequency 499 Hz, DWARF stack depth 16,384, ten
warmups, and 2,000 measured scans per child. The whole process is sampled,
including setup, timer-excluded checks, destruction, and reporting. The
profiled elapsed values are diagnostic evidence and do not replace the 0885
latency baseline.

The separate Callgrind plan uses two fresh runner-scoped processes, ten scans,
and no warmups. Collection starts and stops at
`litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan`. The runner
scope includes oracle preparation and timer-excluded work. Inclusive rows for
MCE processing, MCE `start`, and x14ac capture overlap and must not be added to
one another or combined with sampled native cycles.

The helpers are create-only and fail fast. `capture.py` and
`capture_callgrind.py` retain the command, stdout, stderr, semantic report,
producer evidence, raw profile, and command receipt for each child. A failed
attempt is never retried in place. `export.py` turns retained perf data into
leaf samples, call stacks, flat self-symbols, headers, and per-child
`buildids.txt` views. `reduce.py` reconciles exact periods and sample counts,
then compares disjoint self-symbol shares with 0886. `analyze_callgrind.py`
annotates exact canonical rows and compares guest instruction observations
with 0886. `verify.py` checks source/binary/build identity, frozen commands,
artifact hashes, report binary identity, semantic oracle results, and profile
population custody without launching a workload.

The semantic checks reuse the reviewed 0885 verifier through a private import
adapter. They retain the producer corpus and oracle contract: 4,096 cells,
3,686 numeric values, 410 shared-string values, the archive digest, source
version fence, route/refusal census, and all report vectors.

After the root plan and methodology review are sealed, the intended commands
are:

```sh
python3 -B docs/performance/results/change-0892/capture.py \
  --plan docs/performance/results/change-0892/plan.json \
  --output-root docs/performance/results/change-0892/capture
python3 -B docs/performance/results/change-0892/export.py \
  --capture-root docs/performance/results/change-0892/capture
python3 -B docs/performance/results/change-0892/reduce.py \
  --capture-root docs/performance/results/change-0892/capture \
  --output docs/performance/results/change-0892/self-analysis.json
python3 -B docs/performance/results/change-0892/capture_callgrind.py \
  --plan docs/performance/results/change-0892/callgrind-plan.json \
  --output-root docs/performance/results/change-0892/callgrind
python3 -B docs/performance/results/change-0892/analyze_callgrind.py \
  --capture-root docs/performance/results/change-0892/callgrind \
  --output docs/performance/results/change-0892/callgrind-analysis.json
python3 -B docs/performance/results/change-0892/verify.py \
  --plan docs/performance/results/change-0892/plan.json \
  --capture-root docs/performance/results/change-0892/capture \
  --callgrind-root docs/performance/results/change-0892/callgrind \
  --output docs/performance/results/change-0892/verification.json
```

No source build is performed by these helpers. No profiler output establishes
a wall-clock speedup, allocation improvement, or end-to-end claim. Any final
comparison must name the source, retained executable, corpus, host, protocol,
and metric, and must keep sampled native self attribution separate from
inclusive guest-instruction attribution.
