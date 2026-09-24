# ODS inert DDE catalog performance

This receipt compares the DDE transaction implementation at baseline commit
`ae9bf7bfa11eed50e6560be166eaa18987cda2b5` with a candidate overlay whose only
production change is `crates/litchi-ods/src/model/dde/transaction.rs`.
Baseline transaction SHA256 is `bce59b24064e017d0a3f3cf945594fdd9870e84cdeda513d076ea9c86fc1bc4c`;
candidate transaction SHA256 is
`d8be57303a27c65d68312b18477a9356c5d4df8fa054910dace9b2318f009706`.
The exact baseline and candidate source/harness manifests, executable hashes, and the
production/test replay patch are retained in the adjacent directories.

Both binaries ran the same retained harness and generated the same three
bounded ODF 1.4 `content.xml` shapes: `small` (9,845 bytes, 1 source, 2
links, 8×8 caches), `large-cache` (4,132,098 bytes, 1 source, 1 link,
256×256 cache), and `many-links` (863,375 bytes, 4 sources, 2,048 links,
1×1 caches). DDE topics are inert `file:///never/...` values. No native DDE
producer, ZIP package, or publication path participates in these measurements.

## Method and boundaries

Each of the 18 baseline lanes and 18 candidate lanes used three warmups and
15 measured iterations. The process was pinned to CPU 2; `/usr/bin/time -v`
recorded process maximum RSS. The host was shared with other work, so timing
comparisons are scoped observations and do not establish general ODS
performance. The harness counting allocator reports requested/released bytes
and allocator live-byte deltas for the timed region. `Work` and `Memory` are
execution-budget counters; RSS is a separate whole-process OS value.

`parse` and `read` controls are retained in both raw CSVs. The comparison
lanes below are `stage`, `commit`, and `edit`. Stage constructs the caller's
replacement `LinkSpec` before the timer and times draft staging. Commit parses
and stages before the timer and times transaction rendering and readback. Edit
constructs its typed replacement before the timer and times parse, stage, and
commit; it is a metadata edit workflow, excluding caller spec construction
and package publication. The raw receipts also retain `mean`, `p95`, `p99`,
all allocator counters, budget before/after values, output bytes, and RSS.
The derived rate column is source bytes divided by p50 nanoseconds; it is a
nominal input-byte rate for commit/edit; staging does not scan the input and
has no byte rate shown. This rate does not assert that the implementation copied
that number of bytes. Each pair is `baseline → candidate` and is p50 unless
otherwise stated.

## Critical lane comparison

| lane | source B | p50 ns | nominal source B/s (10⁶) | alloc calls | requested B | allocator peak live Δ B | Work Δ | Memory Δ | max RSS KiB |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| stage-small | 9,845 | 310 → 320 | — | 5 → 5 | 520 → 520 | 520 → 520 | 4 → 4 | 2,657 → 2,657 | 2,336 → 2,252 |
| stage-large-cache | 4,132,098 | 430 → 320 | — | 5 → 5 | 360 → 360 | 360 → 360 | 3 → 3 | 2,103,553 → 2,103,553 | 12,560 → 12,488 |
| stage-many-links | 863,375 | 17,440 → 16,450 | — | 5 → 5 | 327,880 → 327,880 | 327,880 → 327,880 | 2,050 → 2,050 | 327,833 → 327,833 | 4,896 → 4,820 |
| commit-small | 9,845 | 166,381 → 166,381 | 59.171 → 59.171 | 856 → 855 | 132,179 → 131,971 | 35,544 → 35,544 | 21,221 → 21,221 | 11,089 → 11,089 | 2,292 → 2,288 |
| commit-large-cache | 4,132,098 | 73,034,420 → 72,521,835 | 56.577 → 56.977 | 393,381 → 393,381 | 68,176,652 → 68,176,652 | 25,243,341 → 25,243,341 | 9,315,827 → 9,315,827 | 4,133,057 → 4,133,057 | 36,636 → 36,608 |
| commit-many-links | 863,375 | 20,386,187 → 20,593,210 | 42.351 → 41.925 | 112,924 → 110,874 | 448,348,123 → 12,352,299 | 2,360,524 → 2,360,524 | 1,528,777 → 1,528,777 | 1,463,839 → 1,463,839 | 8,328 → 8,692 |
| edit-small | 9,845 | 244,511 → 244,171 | 40.264 → 40.320 | 1,078 → 1,077 | 158,696 → 158,488 | 47,538 → 47,538 | 31,259 → 31,259 | 24,497 → 24,497 | 2,296 → 2,492 |
| edit-large-cache | 4,132,098 | 102,847,557 → 102,229,384 | 40.177 → 40.420 | 458,990 → 458,990 | 76,509,910 → 76,509,910 | 29,377,386 → 29,377,386 | 13,513,999 → 13,513,999 | 10,369,329 → 10,369,329 | 36,652 → 36,608 |
| edit-many-links | 863,375 | 30,904,991 → 30,591,483 | 27.936 → 28.223 | 158,113 → 156,063 | 453,255,631 → 17,259,807 | 3,938,778 → 3,938,778 | 2,412,672 → 2,412,672 | 3,255,173 → 3,255,173 | 8,284 → 8,636 |

The catalog preallocation change has its clearest measured effect in the
`many-links` allocation counter. `commit-many-links` requested bytes fall
from 448,348,123 to 12,352,299 (97.24% lower), and `edit-many-links` from
453,255,631 to 17,259,807 (96.19% lower). Allocator call counts fall from
112,924 to 110,874 and from 158,113 to 156,063 respectively. The
allocator peak-live deltas, Work, Memory, and output bytes remain unchanged
in those rows. This evidence reports allocator requests; it does not claim a
specific number of bytes copied or a latency improvement.

The many-links commit p50 is 20,386,187 ns at baseline and 20,593,210 ns for
the candidate; edit p50 is 30,904,991 ns and 30,591,483 ns. These differences
are small relative to the shared-host timing conditions and should not be
presented as a latency win. The small and large-cache commit/edit allocation
counters are unchanged or differ by one call, as shown in the table. Stage
has the same allocation and budget counters because the catalog scan occurs
at commit, outside the timed staging path.

## Whole-process counter check

`perf stat -r 3` was available. The check used the same CPU pin and
`--warmups 3 --iterations 15` `commit-many-links` command for each binary.
The counters include process startup, parsing, spec preparation, staging, and
all measured iterations; they are not commit-only counters.

| counter (three-run average) | baseline | candidate |
| --- | ---: | ---: |
| cycles | 2,524,659,326 | 2,499,246,023 |
| instructions | 9,066,905,369 | 8,988,222,297 |
| branches | 2,118,541,345 | 2,103,926,894 |
| branch-misses | 2,846,352 | 2,821,503 |
| cache-misses | 1,169,384 | 1,071,085 |
| page-faults | 19,914 | 10,817 |
| elapsed seconds | 0.571599973 ± 0.010118400 | 0.556500904 ± 0.000588152 |

These whole-process counters are supporting observations only. No CPU profile
was collected, so they do not attribute a cost change to a particular stack
or establish causality.

## Reproduction and retained evidence

The complete per-lane receipts are [`baseline/raw.csv`](baseline/raw.csv) and
[`candidate/raw.csv`](candidate/raw.csv), with LF line endings. Baseline and
candidate commands are in [`baseline/commands.txt`](baseline/commands.txt) and
[`candidate/commands.txt`](candidate/commands.txt). The candidate binary is
`f7fbdb56cce346a28c1cf5e45328092b4ba4a8966d2517b7cea16b4cd573942f`; baseline
binary is `39f1bf4796186073dd7e56abccce8986960d633696012380635068bd97ef88b1`.
The candidate 47-test overlay gate is in
[`candidate/scoped-tests-final.log`](candidate/scoped-tests-final.log), and
all 35 transaction plus 12 facade tests passed. Final root ODS gates also
passed; the source was unchanged for their final run.

To replay the production overlay from the exact baseline:

```sh
EVIDENCE=/absolute/path/to/ods-dde-catalog-performance
BASE=/var/tmp/ods-dde-catalog-replay
TARGET=/var/tmp/ods-dde-catalog-replay-target
git worktree add --detach "$BASE" ae9bf7bfa11eed50e6560be166eaa18987cda2b5
git -C "$BASE" apply --unidiff-zero "$EVIDENCE/candidate/transaction.diff"
mkdir -p "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-catalog-performance/candidate/harness/src"
cp "$EVIDENCE/candidate/harness/Cargo.toml" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-catalog-performance/candidate/harness/Cargo.toml"
cp "$EVIDENCE/candidate/harness/Cargo.lock" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-catalog-performance/candidate/harness/Cargo.lock"
cp "$EVIDENCE/candidate/harness/src/main.rs" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-catalog-performance/candidate/harness/src/main.rs"
CARGO_TARGET_DIR="$TARGET" cargo build --manifest-path \
  "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-catalog-performance/candidate/harness/Cargo.toml" \
  --release --locked --offline
taskset -c 2 "$TARGET/release/ods-dde-transaction-profile" \
  --workload commit --case many-links --warmups 3 --iterations 15
git worktree remove --force "$BASE"
```

The repeated comparable counter runs are in
`baseline/perf-stat-commit-many-links.*` and
`candidate/perf-stat-commit-many-links.*`. This batch covers inert metadata
and transaction catalog allocation behavior on the three generated shapes;
it does not cover DDE execution, refresh, native interoperability, package
publication cost, or broad operating-system performance.
