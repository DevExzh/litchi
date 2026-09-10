# XLS User Names ownership profile

Date: 2026-09-10 UTC

This is a focused allocation and latency profile of the pending BIFF8 `User
Names` stream owner. It is evidence for the stream ownership review only. It
does not represent a native XLS corpus, a complete CFB package workflow, or a
general XLS CRUD benchmark.

Two source snapshots were measured: the immutable v5 candidate before the
transaction ownership fixes, and the current worktree after those fixes. The
final source hashes were:

```text
crates/litchi-xls/src/lib.rs          6fb0f1ac31b0c91c9cf63c5dad9d7dbcfdb40005114d6fa6f40b325dbb1a2d39
crates/litchi-xls/src/user_routing.rs 2d60780ce69df404ff272bda48809cbf0c9af99ee8d711cd546112dfea2b24af
user_names/codec.rs       78b9667cf12d7419e47910e572460d60433061ed33f8192eafa7abfa525febe0
user_names/edit.rs        c534a0407e11ec320d91ea4c82375ce5ffd5f5d3a558018d0e96b2004da2e9c2
user_names/model.rs       6576ab3b89e44a1f7fcc2a14bda3037166a5cb23e7840a298e93d1fa6d76119a
user_names/package.rs     9a36a65e1414ed7224ab5db46dd964c0190db6030905d2c9a8311227dbeb9bbf
user_names/tests.rs       a8a9e8c9d93558d7321b5570957453aa464422e1e53021c0c48b5d1be913d30a
workbook/package.rs       bdc61f2175baa873f8c758d76afee66ad6a9d0656805696217f2353aa442a1e9
revision_log.rs           60ad9c4738c1403f62acb74b99ddbc285d9429498a9bb05f9561efacc6add674
```

The pre-fix process runs used the immutable v5 source tree at
`/var/tmp/litchi-xls-user-names-validation-20260910-v5`. The durable
`results/baseline-source.json` records all 16 source, test, and feature-matrix
hashes, while `results/baseline-source.patch` records the complete
baseline-to-final delta. The patch makes baseline replay independent of that
session-retained path.

The harness is a small external Cargo package in [`harness/`](harness/). Run
it from the repository root with:

```sh
cargo run --release --locked --offline \
  --manifest-path docs/report/spec-gap-validation-evidence/xls-user-names/performance/harness/Cargo.toml
```

The dependency is a path dependency on `crates/litchi-xls`; no production
benchmark dependency or source file is added. The harness uses a counting
`System` allocator, warms each operation, then performs 300 iterations for the
one- and 16-user cases and 60 iterations for the 255-user case. Timing is
process-local wall time. `peak_live_delta` is the maximum allocator live-byte
increase during the measured batch relative to the pre-operation baseline.
Five independent process runs were used for the medians below. The machine was
an AMD EPYC 9R45 (32 logical CPUs), Linux under KVM, with rustc 1.95.0 and
Cargo 1.95.0. The standalone harness explicitly enables release LTO and
`panic = "abort"`; it is an allocator-instrumented profile, so its timing is
directional and is not a production latency claim.

To replay the final five-process sample from a clean checkout, build the
durable harness once and invoke its release binary five times:

```sh
final_target=/var/tmp/litchi-xls-user-names-final-replay-target
cargo build --release --locked --offline \
  --manifest-path docs/report/spec-gap-validation-evidence/xls-user-names/performance/harness/Cargo.toml \
  --target-dir "$final_target"
for run in 1 2 3 4 5; do
  "$final_target/release/litchi-xls-profiler-v5" > "final-run-$run.txt"
done
```

To replay the baseline from a final feature commit, create a detached
worktree, reverse the durable source patch, and point a temporary harness
manifest at that reconstructed crate:

```sh
repo=$(git rev-parse --show-toplevel)
feature_commit=$(git -C "$repo" rev-parse HEAD)
baseline_tree=$(mktemp -d /var/tmp/litchi-xls-user-names-baseline.XXXXXX)
git -C "$repo" worktree add --detach "$baseline_tree" "$feature_commit"
git -C "$baseline_tree" apply --reverse --check \
  "$repo/docs/report/spec-gap-validation-evidence/xls-user-names/performance/results/baseline-source.patch"
git -C "$baseline_tree" apply --reverse \
  "$repo/docs/report/spec-gap-validation-evidence/xls-user-names/performance/results/baseline-source.patch"
baseline_harness=$(mktemp -d /var/tmp/litchi-xls-user-names-baseline-harness.XXXXXX)
mkdir -p "$baseline_harness/src"
cp "$repo/docs/report/spec-gap-validation-evidence/xls-user-names/performance/harness/src/main.rs" "$baseline_harness/src/main.rs"
cp "$repo/docs/report/spec-gap-validation-evidence/xls-user-names/performance/harness/Cargo.lock" "$baseline_harness/Cargo.lock"
sed "s#../../../../../../crates/litchi-xls#$baseline_tree/crates/litchi-xls#" \
  "$repo/docs/report/spec-gap-validation-evidence/xls-user-names/performance/harness/Cargo.toml" > "$baseline_harness/Cargo.toml"
baseline_target=$(mktemp -d /var/tmp/litchi-xls-user-names-baseline-target.XXXXXX)
cargo build --release --locked --offline --manifest-path "$baseline_harness/Cargo.toml" --target-dir "$baseline_target"
for run in 1 2 3 4 5; do
  "$baseline_target/release/litchi-xls-profiler-v5" > "baseline-run-$run.txt"
done
```

`feature_commit` is the final feature commit containing the measured source;
the recorded patch and JSON manifest cover all 16 changed or measured crate
paths. Remove the detached worktree and temporary harness/target after the
replay.

For each synthetic stream, `parse_slice` parses a borrowed `Vec<u8>` and must
retain its own source allocation, while `parse_shared` supplies an existing
`Arc<[u8]>`. `edit_create` calls `Snapshot::edit()` and drops the detached
transaction. `noop_commit` creates that transaction and commits without a
change. `rename_commit` stages one name change and commits it.

## Results

Allocation counts and byte totals were stable across all five runs. Median
latency and memory values were:

| users/name bytes | operation | median us/op | allocs/op | allocated bytes/op | peak live delta |
| ---: | --- | ---: | ---: | ---: | ---: |
| 1 / 8 | parse_slice | 0.85 | 7 | 10,188 | 9,600 |
| 1 / 8 | parse_shared | 0.77 | 5 | 9,008 | 9,000 |
| 1 / 8 | edit_create | 0.04 | 0 | 0 | 0 |
| 1 / 8 | noop_commit | 0.10 | 0 | 0 | 0 |
| 1 / 8 | rename_commit | 1.35 | 15 | 11,299 | 9,598 |
| 16 / 16 | parse_slice | 2.97 | 37 | 13,096 | 11,472 |
| 16 / 16 | parse_shared | 2.86 | 35 | 10,344 | 10,088 |
| 16 / 16 | edit_create | 0.04 | 0 | 0 | 0 |
| 16 / 16 | noop_commit | 0.09 | 0 | 0 | 0 |
| 16 / 16 | rename_commit | 3.47 | 45 | 14,199 | 11,470 |
| 255 / 54 | parse_slice | 60.40 | 515 | 97,746 | 60,490 |
| 255 / 54 | parse_shared | 60.45 | 513 | 50,756 | 36,986 |
| 255 / 54 | edit_create | 0.04 | 0 | 0 | 0 |
| 255 / 54 | noop_commit | 0.09 | 0 | 0 | 0 |
| 255 / 54 | rename_commit | 60.76 | 523 | 98,771 | 60,448 |

The 255-user stream is 23,486 bytes. The pre-fix v5 baseline measured
`edit_create` and exact `noop_commit` at 257 allocations and 36,338 allocated
bytes. The final worktree stores the transaction package behind the snapshot's
existing `Arc`, so both operations are allocation-free for all three stream
sizes. This removes the complete semantic-model clone from the ordinary
inspect-or-conditionally-change workflow.

The final rename path transfers the package parsed by `replace_candidate` into
the committed snapshot instead of reparsing it. At 255 users this reduces
rename publication from 1,292 to 523 allocations, from 185,120 to 98,771
allocated bytes, and from 96,678 to 60,448 peak live bytes. Median
allocator-instrumented time falls from 112.34 us/op in the baseline to 60.76
us/op in the final worktree. These are scoped before/after measurements of the
same harness and synthetic source; they are not an end-to-end speedup claim.

The parser still constructs an owned `UsrInfo` string and then clones that
string into `UserEntry`, so each active user has a temporary and retained name
allocation during parsing. This is why the final 255-user shared parse remains
at 513 allocations and 50,756 allocated bytes.

The shared-source seam removes two allocations and 46,990 allocated bytes at
255 users compared with `parse_slice`, and lowers measured peak live bytes by
23,504. This is a useful caller-owned source path, but ordinary `parse` still
has to retain an owned snapshot by contract.

The first two ownership fixes are present in the final worktree and are covered
by the before/after run. The remaining scoped optimization is to decode
`UsrInfo` directly into `UserEntry` or expose a borrowed fixed-layout view so
parsing does not allocate a temporary name only to clone it into the retained
model. Any such change must preserve exact source bytes, fallible bounds, and
the existing immutable snapshot contract.

The complete raw five-run outputs are retained in
[`results/baseline-v5-runs.txt`](results/baseline-v5-runs.txt) and
[`results/final-current-runs.txt`](results/final-current-runs.txt). The
corresponding source, harness, lockfile, executable, and complete baseline
patch hashes are in the adjacent evidence files;
[`baseline-source.patch`](results/baseline-source.patch) and
[`baseline-source.json`](results/baseline-source.json) make the baseline
source replayable without the retained v5 checkout.

The report intentionally makes no throughput, RSS, native Office, CFB package,
or full end-to-end CRUD claim. The allocator counter reports allocator traffic
and live bytes for this process; it is not a resident-set measurement and does
not account for allocator arena retention after an operation.
