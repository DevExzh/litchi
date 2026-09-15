# Evidence: change 0625, a deterministic storage order for the OLE2 writer

Change record:
[`0625-cfb-writer-deterministic-storage-order.md`](../../0625-cfb-writer-deterministic-storage-order.md).

Disposition: **retained; a correctness fix, not an optimization.**
`performance_claim: none`. One file under `crates/` changed —
`crates/litchi-cfb/src/writer/core.rs`, 47 inserted lines and 2 removed — plus
one new test file. It fixes the first of the two adjacent findings change
[0617](../change-0617/README.md) reported and did not fix: `OleWriter::write_to`
iterated a `HashSet` of storage paths to assign directory SIDs, so a document
with two or more storage paths neither of which is a prefix of the other
serialized to different bytes on every run.

Headline numbers: change 0617's eight-storage document produced **64 distinct
files in 64 builds before and 1 after**; over the repository's **214 CFB
fixtures**, 8 were non-deterministic before and **0 after**, while all **203**
that were already deterministic are **byte-identical**, with no length change on
any of the 211 that rebuild.

## Contents

| Path | What it is |
| --- | --- |
| `determinism.txt` | Change 0617's synthetic document — *N* explicitly created sibling storages, one `Payload` stream each — rebuilt 64 times in one process per *N* ∈ {1, 2, 3, 8}, on both legs, with every distinct output digest. Before: 1, 2, 6 and 64 distinct. After: 1 each. Prints both the standard FNV-1a digest and the one 0617's probe used (it multiplied by sixteen times the FNV-1a prime), so the before leg can be matched against `results/change-0617/determinism.txt` digest for digest. |
| `corpus-before.jsonl`, `corpus-after.jsonl` | One JSON line per CFB artifact under `test-data/` (214 of them, found by `is_ole_file`), per leg: sector size, storage and stream counts, whether the storage set contains two incomparable paths, the distinct output lengths and digests over 16 rebuilds, or the typed refusal. |
| `corpus-summary.txt` | The differential between the two, produced by `scripts/compare.py`: the susceptible-fixture table, the count of fixtures that changed (0), the count whose length changed (0), and the storage-count histogram. |
| `instructions.txt` | Callgrind isolation pairs at 2 and 12 rebuilds on two fixtures, three repetitions per leg: +78 Ir per rebuild on a storage-free 984 KB XLS (exactly reproducible on both legs) and +1,493 Ir on a five-storage DOC, which is inside that case's own repetition spread. |
| `cg-summaries.txt` | The raw `summary:` line of every one of the 24 callgrind runs behind `instructions.txt`, so the medians can be recomputed. |
| `test-fails-before.txt` | The new test file run against the unmodified writer at the base commit, with only the test added: three of the six tests fail (64 and 38 distinct outputs in 64 builds, and the declaration-order test), three pass on both legs because they encode what must not change. |
| `gates.txt` | Tails of `cargo fmt --all --check`, `cargo clippy -p litchi-cfb --all-targets --locked`, `cargo doc -p litchi-cfb --no-deps --locked`, `cargo test -p litchi-cfb --locked`, `cargo test` over the three OLE2 format crates and over the five other crates that call `OleWriter::create_storage`, and the head of the corpus differential. |
| `test-formats.txt` | Every `test result` line from `cargo test -p litchi-xls -p litchi-doc -p litchi-ppt --locked`: 145 test binaries, 3,779 passed, 0 failed, 25 ignored. |
| `test-consumers.txt` | The same for `cargo test -p litchi-ole-common -p litchi-crypto -p litchi-vba -p litchi-sign -p litchi-ograph --locked`: 21 test binaries, 312 passed, 0 failed. |
| `probe/` | The scratch probe: `Cargo.toml` (one path dependency on `litchi-cfb`; `<REPO-ROOT>` stands for the checkout the leg was built against) and `src/main.rs`. Three modes: `determinism N R`, `corpus ROOT R` and `rebuild FILE N`. |
| `scripts/capture.sh` | Captures every measurement above for one leg: the determinism runs, the corpus pass and the callgrind isolation pairs. |
| `scripts/compare.py` | Turns the two corpus legs into `corpus-summary.txt`. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`, for the coordinator to merge. |

## Provenance

- Base commit: `1e41983213dc378c13774ed7038c51faf231977f`
  (`perf(ppt): span the retained stream instead of copying every record payload (0606)`).
- Branch: `perf/0625-cfb-writer-deterministic-storage-order`.
- Working copy: a worktree at `/home/zhuhe/code/litchi-worktrees/0625`; the
  shared repository was not built in or modified. The before leg was built from
  the shared read-only checkout at
  `/home/zhuhe/code/litchi-worktrees/before-1e4198321`, which was not modified.
- Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. Other agents
  were building and measuring on the same machine throughout, which is one
  reason this record ranks on deterministic counts and instructions and takes no
  wall-clock measurement at all.
- Toolchain: rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 (f2d3ce0bd
  2026-03-21), valgrind 3.26.0.
- Measured CPU: 19, via `taskset -c 19` on every profiled and counted process.
- Probe build: `--release` with `debug = 1`, own `CARGO_TARGET_DIR` outside the
  repository, at `/home/zhuhe/code/litchi-worktrees/targets/0625-{before,after}`
  (removed after the evidence was copied here).
- Binary sha256:
  - before leg `fac2248590bd418eff4cfa8017ad201edf99231ac9243304e34d4788608de5b1`
  - after leg `947519b3556c059fee10d7ff9d7b4ef9f3e045415e9874e3b194f229263fb821`
- Probe source sha256 (`src/main.rs` and `Cargo.toml` concatenated in that
  order): `ff1387372145ebe32db6f45c763b1429464086ae2d682b1a9836d011b02feb9b`.
- Script sha256: `capture.sh`
  `27432b6bf8ff1eb749c11d198102b8da7213c4243e1b9a31d88ef29a37734d34`,
  `compare.py`
  `17db5505df657ce559a299f2264c1959088960fb14a23acd662065db72cd570f`.

## Reproducing

```sh
# 1. build the probe once per leg, outside the repository
cp -r probe probe-after && sed -i 's|<REPO-ROOT>|<FIXED-CHECKOUT>|' probe-after/Cargo.toml
CARGO_TARGET_DIR=<TARGET-DIR>/after cargo build --release --manifest-path probe-after/Cargo.toml
# ... and the same with <BEFORE-CHECKOUT> into probe-before

# 2. capture one leg (determinism, corpus, callgrind)
BIN=<TARGET-DIR>/after/release/cfb_storage_order_probe REPO=<FIXED-CHECKOUT> \
  OUT=<SCRATCH> LEG=after CPU=19 bash scripts/capture.sh

# 3. difference the two corpus legs
python3 scripts/compare.py <SCRATCH>/corpus-before.jsonl <SCRATCH>/corpus-after.jsonl

# 4. the new tests, and the proof that they fail without the fix
cargo test -p litchi-cfb --test writer_storage_determinism --locked
git stash push crates/litchi-cfb/src/writer/core.rs   # keep only the test file
cargo test -p litchi-cfb --test writer_storage_determinism --locked   # 3 of 6 fail
git stash pop
```

## What is not here

No wall-clock timing and no A/A floor: nothing is claimed about latency, and the
measured instruction cost of the change (+78 Ir on a storage-free save) is three
to four orders of magnitude below this host's A/A timing floor, so a paired
timing could not resolve it. No cycles, allocation, peak-RSS, page-fault,
cold-cache or throughput measurement. No per-symbol attribution: the change is
two allocations and a sort, and the isolation-pair totals bound it. No
digest-comparing differential through each format crate's own editor — the
corpus differential goes through `OleWriter` directly, and the nine consumer
test suites are pass/fail. No measurement of a produced encrypted-OOXML
container, whose two incomparable `\x06DataSpaces` child storages match the
shape the synthetic *N* = 2 case measures but were not separately observed.
