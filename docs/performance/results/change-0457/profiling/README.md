# 0457 bounded large endpoint profiles

This diagnostic samples the normal release process for the large semantic-shape endpoint 100 times after three warmups on CPU 2 with one workload worker. At the retained 0457 timings, that is about 15 seconds per role. It profiles the source-tail candidate and the owned existing-append control separately; the two processes are never run together.

The final retained profiles used these commands serially after the release
gates. The distinct check tags are part of the receipt identity and must not be
reused:

```text
python3 -B docs/performance/results/change-0457/check.py --tag profile-candidate-large-r2 -- python3 -B docs/performance/results/change-0457/profiling/capture-r2.py --role candidate
python3 -B docs/performance/results/change-0457/check.py --tag profile-control-large-r2 -- python3 -B docs/performance/results/change-0457/profiling/capture-r2.py --role control
```

The runner binds the copied normal binaries in `/tmp/litchi-goal-0457/{candidate,control}`, their 0457 workload protocol, binary binding, successful build receipt, source manifest, and independent report oracle. Its workload argv is the existing `--case`, `--semantic-shape large`, `--workers 1`, `--samples 100`, `--warmup 3`, `--json`, and `--corpus-manifest` form. Final output is under `profiling/r2/runs/{candidate,control}-large/`.

The initial candidate attempt stopped before workload execution because four
archived fuzz source inputs had extended the ambient source manifest. R1 binds
those exact additions while requiring every candidate build source to remain
unchanged. It authenticates the older control binary against its own build and
records the current ambient source epoch separately. Before its first run, R1's
JSON writer wrapper was corrected to call the saved original writer.

R1 captured and symbolized the candidate profile but invoked the report oracle
with its default 30-sample expectation. R2 explicitly records and passes
`--samples 100 --warmups 3`; its candidate gate passes. The R2 control workload,
sampling and postprocessing also succeed, but its original control oracle
expects frame-pointer flags. The control's previously accepted `protocol-r1`
and `verify-report-r1.py` already correct that expectation to `rustflags: null`.
The same retained profile report passes that amended oracle:

```text
python3 -B docs/performance/results/change-0457/check.py --tag profile-control-oracle-r1 -- python3 -B docs/performance/results/change-0457/control/oracle/verify-report-r1.py --report docs/performance/results/change-0457/profiling/r2/runs/control-large/report.json --mode normal --shape large --samples 100 --warmups 3
```

`control-oracle-amendment.json` binds that validation to the unchanged report,
original failed profile receipt, applicable control protocol and oracle. All
earlier runner/protocol versions, failures, and raw profile artifacts remain.

`perf record` samples `cycles:u` at 99 Hz. The runner probes `perf record --call-graph help`; it adds `--call-graph dwarf,8192` only when the probe advertises DWARF. It retains `perf.data`, record stdout/stderr, GNU `time -v` resource output, flat `perf report` top symbols, raw `perf script`, and a small local folded-stack conversion. If sampling or symbolization fails, the receipt keeps the failure and whatever raw files exist.

The profile is whole-process diagnostic evidence. `resource.log` retains process-level RSS from GNU `time`; hardware counter totals and ordinary allocator counters are unavailable because this plan does not run `perf stat` or tune an allocator. The oracle runs after recording with `--samples 100 --warmups 3`, outside the sampled process. No API attribution, causal hotspot, speedup/regression, I/O, scaling, or cancellation claim follows from these files.

The candidate profile's largest flat symbols include SHA-256 compression
(11.00%), XML name validation (9.08%), memory comparison (6.01%) and start-element
validation (5.42%). The control shows memory comparison (9.67%), attribute
iteration (6.00%), memory movement (5.76%) and namespace-prefix resolution
(5.24%). These percentages describe the sampled whole process, including setup,
warmups and reporting; they do not isolate an API phase or prove causal savings.
