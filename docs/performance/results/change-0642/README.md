# Change 0642 evidence packet

`SourceBackedWorksheet::visit_cells` produces each cell as it visits it instead
of building a whole-range `Vec<SourceCell>` first, with the order-of-refusal
contract preserved. The record is
[`../../0642-xlsx-visit-cells-streaming.md`](../../0642-xlsx-visit-cells-streaming.md).

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Provenance

| | |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` |
| branch | `perf/0642-xlsx-visit-cells-streaming` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680` (shared, read-only) |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14) — the repository's pinned toolchain — valgrind/callgrind 3.26.0 |
| CPU pin | `taskset -c 17`; seven other agents were building and measuring on the host throughout |
| probe build | `cargo build --release` from `probe/`, once per leg, each with its own `CARGO_TARGET_DIR`; `probe/rust-toolchain.toml` pins the same 1.95.0 the workspace uses |
| harness build | `cargo build --release --locked` in `tools/perf-baseline`, once per leg |
| binary staging | every measured binary was copied out of its Cargo target directory before being run (change 0627) |

Binary identity is in [`probe-sha256.txt`](probe-sha256.txt), which carries the
file `sha256` **and** the `.text`/`.rodata` section `sha256` of each binary,
because two after binaries were rebuilt between a measurement and this table.

* The **probe** was rebuilt after one doc comment in `source.rs` was corrected
  to cite a measurement taken on the pinned toolchain. That is verified, not
  argued: the rebuilt binary's `.text` (`7c592d7a…`) and `.rodata`
  (`52a051b4…`) are byte-identical to the measured one's, and the differential
  and the warm allocation counts were re-run on the rebuild and reproduced
  exactly. Every probe figure in the record is therefore a figure of the
  committed code.
* The **harness** was rebuilt after `cargo fmt` reflowed one expression in
  `cells`. The first harness timing window and the harness callgrind pair used
  the pre-reflow binary; the second window used the staged one. `cargo fmt`
  changes whitespace only, and the two windows agree, but the harness's
  pre-reflow file hash (`73392b60bb30f436baa2906e39de1f4b26026b22ed4346c144c8dcbb58331209`)
  is recorded here rather than left implied.

## Corpora

| corpus | what it is |
| --- | --- |
| `no_drawing_patriarch.xlsx` | `test-data/poi/test-data/spreadsheet/`, 672,414 archive bytes; one worksheet `Лист 1`, 3,382,556 uncompressed worksheet bytes, 75,770 stored cells, 3,440,972-byte shared-string table. This worksheet falls back to the materialized store. |
| `dense-wide-probe.xlsx` | built by `xlsx0642 corpus OUT 2 256 256`; two sheets of 256 × 256 integer cells, 384,231 bytes, sha256 `8591744a…`. It follows the shape of the harness `dense-wide` corpus (`litchi-xlsx-synthetic-v1`, 384,525 bytes) but is not byte-identical to it. This worksheet is eligible for the bounded selected scan. |
| harness `dense-wide` | generated in-process by `litchi-perf-baseline --xlsx-shape dense-wide` for the two range-scan selectors |

## Contents

| path | what it is |
| --- | --- |
| `decision.json` | the disposition, reason codes, accepted costs and known gaps |
| `gates.txt` | the tail of every gate, and the two sets of pre-existing warnings (six in the `litchi` facade, four in `tools/native-resave`) reproduced on the untouched before checkout |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |
| `probe/probe.rs` | the evidence probe: `corpus`, `diff`, `visit`/`cells`, `visit-warm`/`cells-warm` and `bench` modes |
| `probe/alloc.rs` | the counting-allocator companion (allocator shape taken from the change 0634 driver) |
| `probe/Cargo.before.toml`, `probe/Cargo.after.toml`, `probe/rust-toolchain.toml` | the two path-dependency manifests and the pinned toolchain |
| `probe-sha256.txt` | file and `.text`/`.rodata` hashes of every measured binary, plus the probe corpus hash |
| `differential/diff-before.tsv`, `differential/diff-after.tsv` | the four-way differential over 397 worksheets of 182 files, one row per worksheet, one table per leg; `diff -q` between them is clean |
| `differential/tests-on-base.txt` | the three new `streaming_0642_tests` run against the unmodified base implementation |
| `alloc/alloc-raw.txt` | allocation calls, allocated bytes, peak live bytes and retained live bytes for `visit`, `cells`, `visit-warm` and `cells-warm`, both legs, both corpora |
| `callgrind/isolation-pairs.txt` | every isolation-pair total: warm and cold, both corpora, both legs, plus the harness selector |
| `callgrind/inline-experiment.txt` | the `#[inline(always)]` decision, measured both ways on the pinned toolchain |
| `callgrind/harness-nondeterminism.txt` | three isolation pairs of the **same** before harness binary, bounding that selector's count spread |
| `callgrind/cold-cells-symbol-diff.txt` | per-symbol differencing of the cold `cells` pair, which attributes the +0.20% |
| `timing/*.txt` | every paired-timing leg, one elapsed nanosecond count per line |
| `timing/probe-summary.txt` | per-leg p50/mean/p95/p99 and paired deltas in both directions, for every probe scenario and every A/A floor |
| `timing/harness-summary.txt` | the same for the two harness range-scan selectors, with their A/A floors |
| `abba.sh`, `stats.py` | the paired-timing driver and its summarizer |

## Reproducing

The two probe manifests carry absolute `path` dependencies on the worktrees this
batch used, as change 0621's packet does; point them at a checkout of the base
commit and at a checkout of this branch to rebuild the two legs.

```sh
# build both legs of the probe
for leg in before after; do
  cargo build --release --manifest-path probe/Cargo.$leg.toml
done
# the probe corpus
xlsx0642 corpus dense-wide-probe.xlsx 2 256 256
# the differential (both legs must print identical tables)
xlsx0642 diff $(find . -name '*.xlsx') dense-wide-probe.xlsx
# allocations
xlsx0642_alloc visit-warm dense-wide-probe.xlsx Sheet1
# instructions
valgrind --tool=callgrind --cache-sim=no --branch-sim=no xlsx0642 visit-warm dense-wide-probe.xlsx Sheet1 2
valgrind --tool=callgrind --cache-sim=no --branch-sim=no xlsx0642 visit-warm dense-wide-probe.xlsx Sheet1 6
# paired timing
./abba.sh warm-dense visit-warm dense-wide-probe.xlsx Sheet1 5 40 xlsx0642-before xlsx0642-after
python3 stats.py runs warm-dense
```
