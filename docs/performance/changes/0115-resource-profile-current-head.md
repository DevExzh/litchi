# Change 0115: current-HEAD process resource profile

Date: 2026-08-15

Status: Recorded evidence; no production code change.

## Scope and identity

This tranche adds a standard-library-only orchestrator/parser around the
existing `tools/perf-baseline` selectors.  It covers one current-HEAD probe of
OPC source-backed one-Part publication, managed XLSX scalar batch edit/save,
RTF streaming creation, CFB selective MiniFAT/FAT reads, CFB same-length
atomic save, and the existing explicit OPC/CFB scaling selectors at
1/2/4/8/available widths.  iWork is excluded.  It makes no before/after or
optimization claim.

The measured revision is `be500459961471659f65c180de0e5fe98bc14e3a` and the
release binary is SHA-256
`1cbb2340eae13f4ed49d5baa27532e1f9b31d5781036bb2a302837bcd2210f5c` with
36,646,512 bytes.  The aggregate reports a dirty worktree because unrelated
agent edits were present.  The locked release build completed successfully
(`build_content_bound`), so the binary hash/size and captured source/content
identity are bound to this recorded build.  The source identity includes HEAD
tree `739ba8e610208d2528d580595106a88787143098`, status-z SHA-256
`94b0a8c2fdd8f508e18cbb3278b21abea36a535c270cf748e7a81a7fe1cc08ed`, and
head-to-worktree diff SHA-256
`58a78363d20bd4db858f01a96f33735ac418ea0199a010367242780ad90a6f00` over
49,538 bytes; the harness manifest and lockfile are respectively SHA-256
`a86e3cd2aa6a93c5192ad22f958ec25b6cafae02430ee3bd731eb5b987e4d007` and
`14f1246f8ac9c810bfffd8beb3cd7b0feb0ba68c769a0695491b8f4b9de965f7`.  The
recorded build command is:

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
```

The probe command was:

```sh
python3 tools/perf_resource_profile.py run \
  --build \
  --binary tools/perf-baseline/target/release/litchi-perf-baseline \
  --output docs/performance/results/resource-profile-current-head-0115.json \
  --warmup 1 --samples 3 --timeout 600
```

Raw `/usr/bin/time`, heaptrack, perf, strace, and harness reports were kept in
temporary directories only.  The compact result retains each artifact's SHA,
size, command, parsed fields, and corpus/target hashes.

## Environment and tools

The host was Linux `6.8.0-101-generic`, AMD EPYC 9575F, 12 logical CPUs, Rust
`1.95.0 (59807616e 2026-04-14)`, page size 4096, and
`perf_event_paranoid=1`.  The available tools were GNU `/usr/bin/time`,
heaptrack/heaptrack_print 1.5.0, perf 6.8.12, and strace 6.8.  All six
requested perf events returned numeric values in the one-sample probes.

## Observed harness and resource results

The harness ran three measured samples after one warm-up.  Times are the
harness's measured operation interval; `/usr/bin/time` is whole-process RSS.
Heaptrack includes startup, synthetic corpus construction, one measured run,
and profiler overhead.

| Workload / corpus | p50 ns | p95 ns | max RSS KiB | allocation calls | allocated bytes | peak heap | peak RSS with heaptrack |
|---|---:|---:|---:|---:|---:|---:|---:|
| OPC source one-Part / few-large incompressible | 59,684,605 | 59,822,185 | 118,176 | 1,576 | 306,633,284 | 132,791,664 | 126,573,608 |
| Managed XLSX batch / cell-values medium | 33,260,724 | 33,895,459 | 66,132 | 6,130,956 | 1,026,348,498 | 63,239,618 | 75,801,559 |
| RTF streaming / medium | 10,016,573 | 10,114,007 | 30,080 | 450,852 | 66,379,667 | 26,025,656 | 35,232,153 |
| CFB selective MiniFAT / 36-byte target | 140,654 | 145,330 | 30,336 | 13,589 | 148,580,902 | 23,142,072 | 27,682,406 |
| CFB selective FAT / 4 MiB target | 374,947 | 1,225,272 | same paired process profile | 13,589 | 148,580,902 | 23,142,072 | 27,682,406 |
| CFB same-length atomic save / few-large | 156,307,917 | 157,041,972 | 110,884 | 1,722 | 460,627,078 | 115,186,073 | 122,704,363 |

Logical counters are intentionally separate from physical observations.  The
OPC source case recorded 549 source reads / 16,785,201 source bytes per sample,
one ordinary payload materialization, and a 16,783,632-byte sink with 461
writes.  Managed XLSX recorded 225 reads / 4,230,793 bytes, six
materializations, and a 4,226,645-byte sink with 163 writes.  RTF retained zero
output bytes and a 37-byte authoring window while accepting 630,819 bytes in
90,122 writes.  CFB selective returned 36 bytes from the MiniFAT target and
4,194,304 bytes from the FAT target.  CFB save reported 1,825 logical reads /
84,838,500 bytes per sample, one changed span, and a 16,913,408-byte
publication.

The external syscall trace is whole-process `read`/`write` evidence.  For
example, one-sample strace returned 203 read calls / 82,897 bytes and 1,279
write calls / 4,777 bytes for OPC source publication; the CFB save wrapper
returned 1,130 read calls / 122,851,854 bytes and 16,885 write calls /
79,904,360 bytes.  These values do not measure decompression, recompression,
memory copies, or physical storage I/O.

The explicit scaling selectors used the existing bounded execution context.
On the many-small incompressible corpus, OPC p50 was 567,473 ns at one worker
and 789,610 ns at 12 workers; CFB p50 was 224,090 ns and 225,201 ns.  Both
were classified `nonideal_or_measurement_noise`: the raw p50s showed no
measured speedup, and at least one derived Amdahl fraction was outside the
physical [0,1] range.  Per-width speedup, efficiency, and the raw fraction are
retained in the JSON, while invalid fractions are represented as a null
estimate with an explicit validity flag rather than claimed as a
serial-fraction result.

## Limitations

This is a modest warm-memory synthetic-corpus probe.  It does not establish a
cold-cache result, remote/range behavior, production allocation attribution,
confidence intervals for external one-sample profiles, decompressed or
recompressed bytes, memory-copy volume, or a before/after change.  No source
or sink implementation was changed, and no optimization is accepted from this
record alone.
