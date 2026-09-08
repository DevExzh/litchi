# 0466: current dense XLSX commit/save profile

This descriptive bundle investigates the next default-baseline latency lead.
It changes no production/harness Rust or checked coverage. See the
[per-change record](../../changes/0466-xlsx-dense-commit-profile.md),
[source path](source-review.md), and [next implementation](next-work.md).

The normal executable is the exact 0465 build, authenticated by adjacent
`../change-0465/binding.json`, its build receipts and 7,032-file source manifest.
All those sources and all reviewed ADR files were checked unchanged before
and after capture. The runtime reports' Rust 1.95.0/current-revision probe is
not build provenance: the reused binary was built with Rust 1.98.1 during
0465. The separate `profile-binding.json`/`profile-build.json` record the
same-source Rust 1.98.1 diagnostic build with debug level 1, forced frame
pointers and unwind tables. Profile-build timing is not normal latency evidence.

| Artifacts | Scope |
|---|---|
| `normal-r1`–`normal-r4` | Four fresh-process normal runs, each 30 samples/3 warmups, CPU 2, one worker |
| `counters` | Whole-process perf-stat counters, 30/3 |
| `samples` and `summary-dwarf.json` | Original normal-binary DWARF profile, 50/3; incomplete callchains remain explicit |
| `samples-fp` and `summary-fp.json` | Same-source diagnostic frame-pointer profile, 50/3; primary CPU attribution |
| `heaptrack` | Whole-process allocation stacks, 5/1; raw compressed trace and print export |
| `compression.json` | Original and gzip identities, verified equal after decompression before deleting originals |
| `host.json`, `adr-refresh.json` | Environment and unchanged ADR review binding |

The whole-process profilers include fixture generation, expected-output
construction, warmup, reopen/verification, report creation and teardown.
The harness timer contains only ordinary commit plus sequential write after
source open, edit planning and sink reservation. Ancestor weights are
inclusive CPU samples, not elapsed phases. The exact commit marker includes
untimed expected-output construction when its stack is retained. Inclusive
rows overlap. The explicit `CountingSink` writer marker does not classify
other package writers as the measured output path.

`overlap-review.json` discloses two team-owned `perf report` processes observed
during the initial sequence. Both were terminated and verified absent before
supplemental normal R3/R4. Initial reports are retained; no cause or speedup
is inferred from their difference. The sampled workload completed before a
wrapper `stat.st_size` serialization error. Its recovered receipt records
recovery time and a missing original start timestamp; `capture-initial.py.txt`
preserves the faulty helper. No sampled workload was rerun to repair that log.

`capture.py` reproduces the seven normal-build lanes one at a time in a fresh
output bundle, using an exact authenticated cached executable.
`capture-r1.py.txt` is the helper version used before adding supplemental
lane names. `profile_capture.py` builds, binds, captures, and exports the
separate frame-pointer diagnostic. `postprocess.py` records the initial
profile/Heaptrack exports with debuginfod disabled. Their retained argv arrays
contain the exact commands. Re-symbolization requires the matching executable;
the published text exports replay independently of executable availability.

From a checkout containing this bundle and adjacent 0465 provenance:

```sh
python3 -B docs/performance/results/change-0466/verify.py
python3 -B docs/performance/results/change-0466/test_analyze.py
python3 -B docs/performance/results/change-0466/tests_verify.py
```

`analyze.py` reads retained text exports without invoking perf. Its summaries
bind decompressed inputs using bundle-relative identities. The verifier checks
reports and corpus identities, statistical recomputation, capture provenance,
compressed raw artifacts and CPU summary recomputation. Tests include semantic
tampering after updating receipts/seals, not merely unmodified hash checks.
The checksum seal authenticates every regular bundle file except itself.

The 37-case/201-row default matrix and 11 measured/22 correctness-only CRUD
mappings remain unchanged. No independent producer, native application,
bounded-streaming, physical-cold, remote-source or scaling claim follows from
this owned-byte, single-worker synthetic investigation.
