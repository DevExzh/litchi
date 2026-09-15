# Evidence: change 0579, resume the CFB allocation-chain walk

Change record:
[`0579-cfb-resumable-chain-walk.md`](../../0579-cfb-resumable-chain-walk.md).

Disposition: retained. `performance_claim: none`. This packet carries paired
medians in two directions with a measured noise floor, isolated hardware
counters, per-open chain-step counts and a corpus-wide I/O differential — not a
registry claim.

## Contents

| Path | What it is |
| --- | --- |
| `summary.json` | The folded document every table in the record is read from: chain steps, callgrind totals, `perf stat` rows, both A/B/B/A windows, the counter cells and the corpus differential's digests. |
| `replay.py` | Recomputes and re-checks every cited number from this directory alone. Runs no repository code and needs no build. |
| `chain-steps.json` | Per-open `next_chain_sector` calls and whole-open Ir, folded from threshold-100 callgrind isolation pairs, for eight fixtures by two legs. |
| `counters/summary.json` | The 18 deterministic counter cells: three source implementations by three operations by two legs, 20 warmups and 100 samples each. For every cell it records the **distinct** logical-counter tuples and how many samples each covers, so "constant across all 100 samples" is checkable rather than asserted. |
| `perf/<leg>-<stem>-<mode>-open-s{100,1100}.csv` | `perf stat -x,` isolation pairs. Differencing an 1,100-sample and a 100-sample child and dividing by 1,000 isolates one operation — change 0564's method, reused unchanged by changes 0574 and 0576. |
| `callgrind/ann-<leg>-<stem>-s<n>.txt` | `callgrind_annotate` self cost for each isolation pair. |
| `callgrind/incl-<leg>-<stem>-s<n>.txt` | The same pairs with `--inclusive=yes`. |
| `callgrind/chain-edges.txt` | A compact extract of every `next_chain_sector`, `read_stream_range`, `read_minifat_range` and `GlobalsBuffer::ensure` edge block, taken from `--threshold=100 --tree=caller` annotations so no small edge is dropped. `replay.py` recomputes every per-open chain-step figure from this file. |
| `abba/window-{1,2}.json` | The two paired A/B/B/A latency matrices, folded: per-leg p50/p95/p99, the two directions, the same-binary floor, and whether the four legs of each cell agreed counter for counter. |
| `corpus.{before,after}.txt` | One source-backed open of every `.xls`/`.xlt` fixture under `test-data` through a range-recording positional source: read count, read bytes, source-version count, an FNV-1a digest of the exact positional read ranges, and the refusal text for the seven that are refused. The two files are byte-identical. |
| `scripts/` | Every driver used, including `make_corpus_probe.py`, which injects the temporary corpus probe. |
| `environment.json` | Host, toolchain, both tree definitions and both binary hashes. |
| `decision.json` | The `litchi-perf-change-decision` record. |

## Provenance

Base revision `32d25e08806d93f792ffd4954d83acc9db9c5301`. Both legs were built
from a **detached git worktree of that revision outside the repository working
copy**, each with its own `CARGO_TARGET_DIR`, so no file another agent was
editing could enter a measured binary. The after leg differs from the before leg
only in `crates/litchi-cfb/src/shared.rs`, `crates/litchi-cfb/src/lib.rs` and
`crates/litchi-xls/src/workbook/source.rs`. Binary hashes are in
`environment.json` and in every `counters/` cell.

## Result

A source-backed open of `ConditionalFormattingSamples.xls` walked **5,796** FAT
chain links and now walks **2,099**. Instructions per open fall **11.44%**,
cycles **8.34%**, and the paired median falls **8.5%**, against a same-binary
noise floor of **0.94%**. **Not one byte of I/O moved**: reads, read bytes,
source-version observations, the exact positional read ranges and the refusal
texts are identical over all 126 XLS fixtures.

## Replay

```sh
python3 -B docs/performance/results/change-0579/replay.py
```

Rebuilding the captures needs the attribution binary from a worktree of the base
revision:

```sh
cd tools/perf-baseline && cargo build --release --locked \
  --features xls-source-attribution --bin xls_source_attribution
docs/performance/results/change-0579/scripts/capture_counters.sh   <binary> <out-dir>
docs/performance/results/change-0579/scripts/capture_perf.sh       <binary> <out-dir> <leg>
docs/performance/results/change-0579/scripts/capture_callgrind.sh  <binary> <out-dir> <leg>
docs/performance/results/change-0579/scripts/capture_abba.sh       <before> <after> <out-dir>
```

`capture_counters.sh` is change 0574's retained driver with one line changed: the
scratch directory it exports as `TMPDIR`. Everything else about how the children
are launched — CPU 17 pinned, ASLR disabled, single-threaded, 20 warmups and 100
samples — is byte-identical to that record's, which is what makes this packet's
before leg directly comparable with 0574's figures.

The corpus differential is rebuilt by injecting the probe into a worktree and
running it once per leg:

```sh
python3 -B docs/performance/results/change-0579/scripts/make_corpus_probe.py \
  <worktree>/crates/litchi-xls/tests/source_backed.rs
cargo test -p litchi-xls --test source_backed zz_corpus_globals_reads -- --nocapture
```

## What is not here

The raw per-sample A/B/B/A captures — 80 MB of elapsed samples across two windows
— were folded into `abba/window-{1,2}.json` and discarded; the folded documents
keep every quantile and every counter tuple the record cites. The raw callgrind
`.out` profiles and their full `--threshold=100` annotations, 20 MB, were reduced
to the threshold-99 self and inclusive reports plus `chain-edges.txt`, from which
`replay.py` recomputes every chain-step figure. The raw 100-sample counter
captures were folded into `counters/summary.json`, which keeps the distinct
counter tuples and their sample counts rather than 1,800 individual records.

No cold-cache, physical-device, remote or range-source, peak-RSS,
allocation-profile, concurrency-scaling, real-producer or cross-platform result
is here, and no DOC, PPT, XLSX or non-OLE2 measurement.

The rejected first implementation — a retained `SharedOleStreamCursor` — is
described in the record and is **not** retained here as code or as a capture
directory; its measured figures are quoted in the record from captures taken in
the same session and then discarded, because the binary they describe no longer
exists and is not reproducible from this packet.
