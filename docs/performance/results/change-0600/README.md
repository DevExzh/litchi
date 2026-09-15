# Evidence packet — change 0600

One fewer source observation per cold OPC Part read, a bounded monitored-read
scope, and an allocation-free ZIP member lookup. Record:
[`docs/performance/0600-opc-cold-read-observations-and-name-lookup.md`](../../0600-opc-cold-read-observations-and-name-lookup.md).
`performance_claim: none`.

## Provenance

| field | value |
| --- | --- |
| base commit | `8fe9efa55728cf8b9592934f9e9de6f5a66fdc1b` (`feat/office-format-completeness`) |
| branch | `perf/0600-opc-cold-read-observations-and-name-lookup` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-8fe9efa55` (shared, read-only) |
| host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, valgrind 3.26.0 |
| build | `cargo build --release --locked` on both legs |
| CPU affinity | every measured process pinned to CPU 22 with `taskset` |
| host state | eight agents building and testing concurrently throughout; see *Timing* for what that costs |

Binary SHA-256 (`binary-sha256.txt` carries the same list):

| binary | before | after |
| --- | --- | --- |
| `litchi-perf-baseline` | `d0d8f34e…76423f` | `cbfe440a…4d124` |
| `zip-opc-read-probe` | `3b35411e…28bf52` | `462fc815…6a742` |
| `corpus_digest` | `fa2db49b…534771` | `fa2caf06…61e2e` |
| `open_part_loop` | `b5fe09d4…5c57e` | `5197ebb4…2bb4ff` |
| `file_source_time` | `bad925b8…f52cb6` | `f31e590f…aef17` |

## Contents

| path | what it is |
| --- | --- |
| `counts/counts-before.txt`, `counts/counts-after.txt` | the probe's per-phase source observations, allocations, allocated bytes, positional requests and requested bytes on three real OOXML fixtures, one leg each. `diff` of the `requests=` lines is empty. |
| `callgrind/{,pptx-,docx-}{before,after}-{1,11}.log` | valgrind summaries for the N = 1 and N = 11 isolation pairs of `open_part_loop` on the three fixtures |
| `callgrind/full-{before,after}-11.txt` | `callgrind_annotate --threshold=99.5` self-cost tables for the N = 11 xlsx runs; the per-symbol table in the record is their difference |
| `differential/corpus-digest.txt` | the corpus digest: 179 packages, 3,450 lines, every Part's name, content type, `data()` digest, `stream_to()` digest, a second `data()` taken after that stream, every relationship edge, and three refusal probes per package |
| `differential/corpus-digest-sha256.txt` | sha256 of the before leg run twice and of the after leg run twice (once per after binary). All four are equal, which is both the differential result and the proof that the digest is deterministic; only one copy of the 948 KB file is retained because all four are byte-identical. |
| `probe/` | the complete source of every probe binary, with `Cargo.toml`. Path dependencies point at the working copy; retarget them per leg. |
| `timing/abba.sh`, `timing/probe-abba.sh` | the paired-timing drivers, order A1 B1 B2 A2 |
| `timing/abba-*.json`, `timing/probe-time-*.txt` | the raw timing legs |
| `timing/summary-*.txt` | the summarized paired deltas in both directions with the A/A floor |
| `gates.txt` | the tail and aggregate of every gate, plus the differential verdict |
| `binary-sha256.txt` | full SHA-256 for every measured binary on both legs |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |
| `decision.json` | the decision record |

## How to replay

```sh
# counts (three fixtures, one leg)
taskset -c 22 <leg>/zip-opc-read-probe \
  test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx /xl/worksheets/sheet1.xml

# instructions (isolation pair; difference the two and divide by 10)
for n in 1 11; do
  taskset -c 22 valgrind --tool=callgrind --cache-sim=no --branch-sim=no \
    --callgrind-out-file=cg-$n.out <leg>/open_part_loop $n \
    test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx /xl/worksheets/sheet1.xml
done

# corpus differential (run on both legs and diff)
taskset -c 22 <leg>/corpus_digest test-data/ooxml > digest-<leg>.txt

# paired timing
timing/abba.sh <outdir> 3 30 r1
timing/probe-abba.sh <outdir> test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx \
  /xl/worksheets/sheet1.xml reads_after_stream xlsx-after-stream 10 40
```

## What this packet does not contain

No cold-cache, physical-device, remote or range-source, peak-RSS,
concurrency-scaling, real-producer or cross-platform measurement. No claim
registry entry. The timing legs are reported beside their floor and support no
speedup claim.
