# Change 0627 evidence packet — OLE2 range-source selectors

Record: [`docs/performance/0627-ole2-range-source-selectors.md`](../../0627-ole2-range-source-selectors.md).
`performance_claim: none`. No file under `crates/` was modified.

## Contents

| Path | What it is |
| --- | --- |
| `scripts/baseline.sh` | The whole capture: twelve legs, one staged binary, CPU 20, sequential. |
| `scripts/summarize.py` | Produces the record's three tables from `raw/`, and fails if the two transports disagree on logical read calls, logical read bytes or the observation for any fixture/scenario pair. |
| `raw/owned-<fixture>.json` | Owned-source control legs, schema-1 reports (A1). |
| `raw/aa-owned-<fixture>.json` | The same four invocations repeated back to back in the same window (A2), which is the A/A floor. |
| `raw/range-<fixture>.json` | Range-source legs over `SimulatedRangeSource`. |
| `gates.txt` | The tail of every gate that was run, the pre-existing failures with their reasons, and the capture log. |
| `decision.json` | The machine-readable decision record. |
| `log-sections.md` | The four log paragraphs for HOTSPOTS, GOAL_AUDIT, REPORT and ADR_COMPLIANCE, which the coordinator merges. |

`raw/range-54016.json` and `raw/range-withcustomviews.json` are retained
gzipped: uncompressed they are 28.8 MB and 1.9 MB, because
`source.simulation.physical_request_sizes` holds every request size for every
one of the 50 retained samples (16,145 of them per sample on `54016.xls`).
`scripts/summarize.py` reads either form, and the summary it produces from the
gzipped files is byte-identical to the one produced from the plain files.

## Provenance

| Field | Value |
| --- | --- |
| Base commit | `344ed0298fdabe42c2596a4309db8c7a11eaf355` |
| Branch | `perf/0627-ole2-range-source-selectors` |
| Worktree | `/home/zhuhe/code/litchi-worktrees/0627` (detached from the shared working copy) |
| Binary | `litchi-perf-baseline`, `cargo build --release --locked`, SHA-256 `d40979be2b58a1823b28dc4d12e144e36a9862171a1dfd8ce2820766b8728fd1` |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14) |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| CPU pin | `taskset -c 20` for every measured process |
| Window | 18:36:52–19:20:28 UTC, one continuous run, seven other agents active on the host |
| Iterations | 20 warm-ups, 50 retained samples per case |

The binary is staged outside any Cargo target directory before the run. A first
capture was discarded because a concurrent `cargo test --release` relinked
`target/release/litchi-perf-baseline` in place and the in-flight leg died with
`ENOENT`; `scripts/baseline.sh` now takes the staged path as an argument and
says why in a comment. Rebuilding from the same source reproduced the same
SHA-256.

## Transport

Change 0572's delayed arm, expressed in this harness's four-parameter model:

```
--range-fixed-latency-us 1000        1 ms of fixed service per request
--range-request-overhead-us 0        0572's model has no overhead term
--range-bandwidth-bytes-per-sec 104857600   100 MiB/s
--range-max-physical-bytes 65536     64 KiB maximum physical range
```

One model difference from 0572, stated in the record: 0572's probe serves a
capped request as a short read and lets the caller loop; this simulator loops
the cap inside one `read_at`. Request counts agree; call boundaries do not.

## Fixtures

| Fixture | Bytes | Archive SHA-256 | CFB streams | Target stream |
| --- | ---: | --- | ---: | --- |
| `test-data/ole/xls/WithCustomViews.xls` | 165,888 | `3c0c168f38498cc7a356ffaee82b19241ab022ca8813c4ebef286399c32cbd64` | 4 | `Workbook` |
| `test-data/ole/xls/ConditionalFormattingSamples.xls` | 1,402,368 | `d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7` | 8 | `Workbook` |
| `test-data/poi/test-data/spreadsheet/54016.xls` | 984,576 | `2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a` | 4 | `Workbook` |
| `test-data/poi/test-data/slideshow/45543.ppt` | 385,024 | `218aaac542e5f9b567736407f2631defc65797c6ba2a7818f066e2f93bcfacaf` | 5 | `PowerPoint Document` |

The fixture is a caller-named file, not a repository-fixed corpus, so its
identity is recorded per result rather than in
`docs/performance/CORPUS_MANIFEST_V2.md`.

## Reproducing

```sh
git -C <repo> switch --detach 344ed0298fdabe42c2596a4309db8c7a11eaf355   # then apply this branch
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
mkdir -p /tmp/0627-bin && cp tools/perf-baseline/target/release/litchi-perf-baseline /tmp/0627-bin/
bash docs/performance/results/change-0627/scripts/baseline.sh \
  "$PWD" "$PWD/docs/performance/results/change-0627" 20 /tmp/0627-bin/litchi-perf-baseline
python3 docs/performance/results/change-0627/scripts/summarize.py \
  docs/performance/results/change-0627
```

Do not run any Cargo command against `tools/perf-baseline` while the capture is
in flight.

## What is not here

No cold-cache, filesystem, physical-device, network or cross-platform capture.
No allocation, peak-RSS, instruction-count or syscall measurement. No ABBA,
because this record proposes no candidate to compare. No DOC and no CFB-level
selector: change 0587's gap 4 names four formats and this packet covers two.
One transport point only.
