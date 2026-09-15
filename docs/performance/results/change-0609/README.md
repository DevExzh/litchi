# change-0609 evidence packet: the facade `.doc` route, sized against the source-backed DOC reader

Change record:
[`docs/performance/0609-facade-doc-source-route-design.md`](../../0609-facade-doc-source-route-design.md).

Disposition: **retained, design only.** `performance_claim: none`. **No file
under `crates/` was modified.** Everything here compares two routes that both
already exist at the base commit:

- **route E** — `litchi::Document::open(path)` for `.doc`: one whole-file read
  into a `Vec<u8>`, then the eager `doc::Package` parse;
- **route S** — `litchi_doc::body_text::source::SourceSnapshot::open` over a
  `litchi_core::FileSource`.

## Provenance

| | |
| --- | --- |
| Base commit | `8fe9efa55` (`perf(zip,opc): share one Deflate decoder across an OOXML open's structural reads`) |
| Branch | `perf/0609-facade-doc-source-route-design` |
| Worktree | `/home/zhuhe/code/litchi-worktrees/0609` (deleted after this packet was assembled) |
| Before checkout (used only for the pre-existing-warning gate) | `/home/zhuhe/code/litchi-worktrees/before-8fe9efa55` (read-only, detached at `8fe9efa55`) |
| Host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), `--release` |
| valgrind | 3.26.0 |
| CPU pin | **29**, via `taskset -c 29`, on every measured process |
| Concurrency | eight measurement agents active on the host throughout |

There is no before/after pair in this packet, because nothing changed. The
comparison is between two routes over one tree.

### Probe binaries (sha256)

The probe was built twice. **Pair A** produced the callgrind, `perf stat` and
paired-timing numbers. **Pair B** is pair A with one `clippy::while_let_loop`
style fix inside the `identity` subcommand — a mode no timed or profiled run
reaches — and is the source retained in `probe/`. Pair B reproduces **every**
deterministic output of pair A byte-for-byte (`census.tsv`, `identity.tsv`,
`alloc.txt`, `readat.txt` were re-run and diffed clean), and a paired-timing
spot check on `noheadfoot-litchi.doc` reproduces pair A within the A/A floor:
route E p50 26,713 ns against 26,160 (+2.1%), route S p50 167,466 against
164,779 (+1.6%), ratio 6.27× against 6.30×
(`measurements/bench/noheadfoot-litchi-pairB-*.txt`).

| pair | sha256 | produced |
| --- | --- | --- |
| A | `8d956d66b72e88536cddc5406654cf220b91622fd56b6862923ca82f36f539aa` | `callgrind.tsv`, `callgrind-annotate.txt`, `perfstat.tsv`, `strace-summary.txt`, `bench/*-{A1,A2,B1,B2,AA1,AA2}.txt` |
| B | `080dd424d92c90e8a5806161af370c565c043e1eb3987de83830c8bbfe7d34f5` | `census.tsv`, `identity.tsv`, `alloc.txt`, `readat.txt`, `bench/*-pairB-*.txt` |

Build directory (deleted after this packet was assembled):
`/home/zhuhe/code/litchi-worktrees/targets/0609-probe`; the gate-5 check
directory `/home/zhuhe/code/litchi-worktrees/targets/0609-before` likewise.

## Contents

| path | what it is |
| --- | --- |
| `probe/` | the retained scratch driver (`Cargo.toml` with path dependencies, `rust-toolchain.toml`, `src/main.rs`); subcommands `census`, `identity`, `alloc`, `readat`, `loop`, `bench`. Its `Cargo.lock` is not retained: `*.lock` is gitignored in this repository, as it is for change 0587's probes |
| `measurements/census.tsv` | admission of **all 57** `.doc` fixtures under `test-data/` through both routes, with the error kind of every refusal |
| `measurements/identity.tsv` | value comparison on all 57: route E's `text()` length and digest, `paragraph_count()`, `paragraph_text(0)` length and digest; route S's `paragraph(0)` outcome, length and digest, and the length and digest of every position it serves before refusing |
| `measurements/alloc.txt` | counting-global-allocator peak, retained, total and allocation count for six operations on eleven fixtures |
| `measurements/readat.txt` | counting `ReadAt` adapter: `read_at` calls, bytes, `version()` and `len()` calls per route-S operation |
| `measurements/strace-summary.txt` | whole-process syscall isolation pairs (1 against 11 operations, difference divided by 10) for both routes |
| `measurements/callgrind.tsv` | callgrind isolation pairs; `lo_n`/`hi_n` are the operation counts, `ir_per_op` the difference divided by the operation delta |
| `measurements/callgrind-annotate.txt` | per-symbol tops behind those totals, including route S's 87.95% / 87.55% `sha2::sha256::soft::unroll::compress` share |
| `measurements/perfstat.tsv` | native `perf stat -r 3` isolation pairs: cycles, instructions and IPC per operation |
| `measurements/bench/` | paired wall-clock timing, one nanoseconds-per-operation line per sample; `A*` route E, `B*` route S, `AA*` the A/A control, `*-pairB-*` the cross-binary spot check |
| `measurements/fixtures.txt` | the eleven-fixture short-name to path map the scripts use |
| `gates.txt` | every gate's tail |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |

## How to reproduce

```sh
cd docs/performance/results/change-0609/probe
# only if your checkout is not /home/zhuhe/code/litchi:
# sed -i "s@/home/zhuhe/code/litchi@<your checkout>@g" Cargo.toml
CARGO_TARGET_DIR=<scratch> cargo build --release
cd <your checkout>
B=<scratch>/release/facade_doc_route

# deterministic
taskset -c N $B census   test-data          # admission crosstab
taskset -c N $B identity test-data          # value comparison
taskset -c N $B alloc  facade-open <fixture>
taskset -c N $B readat snap-open   <fixture>

# profiled and timed
taskset -c N valgrind --tool=callgrind --cache-sim=no --branch-sim=no \
  $B loop snap-open <fixture> 0 2     # and 0 12; difference / 10
taskset -c N perf stat -x, -e instructions,cycles -r 3 \
  $B loop facade-open <fixture> 10 200   # and 10 1200; difference / 1000
taskset -c N $B bench facade-para0 <fixture> 50 32 40
```

Modes: `facade-open`, `facade-text`, `facade-count`, `facade-para0`,
`snap-open`, `snap-para0`, `snap-open-try` (a route-S open with the refusal
swallowed — the probe cost R-A would add before falling back).

## The five results the record rests on

1. **Admission.** 57 fixtures: route E admits 42, route S admits 8, both admit
   4, and **route S admits 4 that route E refuses** — all four with
   `CorruptedFile("invalid stylesheet: style names and aliases must be
   unique")`, a validation route S never performs.
2. **Capability.** Of the 8 route S admits, 4 serve `paragraph(0)`; of those, 2
   are also admitted by route E, and on those 2 the paragraph text is
   byte-identical. Route S has no `text()` or `paragraph_count()` primitive, and
   walking every position it serves reaches 31 of `noheadfoot-litchi.doc`'s 134
   text bytes.
3. **Reads.** Route S reads the complete artifact **six times** per open (two
   per identity pass, three passes), 6.05-7.47× the file; route E reads it once,
   in 2 `pread64` against route S's 30, with 11 `statx` against 153.
4. **Cost.** Route S is 2.12-5.07× route E's native cycles, 7.9-15.7× its
   instructions and 6.30-9.90× its p50 latency on the one query both answer
   identically, against an A/A floor of p50 ≤0.9% and p99 ≤1.9%.
5. **Memory.** Route S's allocator peak is 47.9-83.8% lower and its retained
   bytes 80.1-97.3% lower. This is the only axis on which it wins.
