# 0362ye fixed admission probe

This is reduced allocator evidence for the fixed-buffer ODP admission probe.
The bounded writer-memo candidate is source commit `9834f169b1`. The generated
archive is the tracked `odp_0362ye_probe` fixture (2,000 slides, 4 MiB media).

| route | e4a baseline peak / allocated | current peak / allocated | output bytes | output SipHash |
|---|---:|---:|---:|---:|
| `open` | 693231 / 777196 | 83233 / 167605 | 0 | 15130871412783076140 |
| `first_slide` | 693231 / 3210784 | 89988 / 2681513 | 47 | 15128627771916252792 |
| `first_text` | 1500637 / 4672388 | 990939 / 4143117 | 105779 | 7807713202416953651 |
| `write_text` | 693231 / 2350086 | 694399 / 2432088 | 105778 | 18333823960124859280 |
| `repeat_write_text` | 693231 / 3922976 | 694399 / 4005356 | 211556 | 6976564693146433600 |
| `unchanged_write` | 6074765 / 6782817 | 5484747 / 6173226 | 4215428 | 12221861870545843138 |
| `append_commit_write` | 612845 / 5737255 | 612893 / 5853815 | 4215645 | 5546059324926701325 |
| `replay` | 612845 / 5737255 | 612893 / 5853831 | 4215645 | 5546059324926701325 |

The earlier fixed-buffer candidate binary is kept outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362yd-candidate-target/release/examples/odp_0362ye_probe`
(3,599,992 bytes, SHA-256
`032047b9f559c291cc07b1052e5425b688a2e91784220c7ef57672d091136bf0`).
The final source build used for the current Callgrind and allocator rows is
`/home/zhuhe/code/litchi-worktrees/uo2-odp-heap-repair/target/release/examples/odp_0362ye_probe`
(3,600,944 bytes, SHA-256
`8f9260bd53df7e615cef498de21ed4124af205c1e4cfe7bc98f3c631eed0e3fc`).
The baseline binary is
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362yd-baseline-target/release/examples/odp_0362ye_probe`
(3,096,104 bytes, SHA-256
`ecb034b0f57588c7d131ba9981d7004f5181ee214a88d14519f73660f35e0669`).

The candidate was built and run with:

```text
ODP_PROBE_ROUTE=docs/unified-ops/results/0362ye/route_after.rs \
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
CARGO_TARGET_DIR=$CANDIDATE_TARGET TMPDIR=$ODP_TMP \
cargo build --offline -p litchi-odp --example odp_0362ye_probe --release
$CANDIDATE_TARGET/release/examples/odp_0362ye_probe open
```

The baseline used the same command with `route_before.rs` and its protected
target. This is allocation evidence; the route-local Callgrind comparison
below uses the same fixed baseline binary and the final bounded-writer candidate.

## Route-local Callgrind

The fixed baseline rows are from
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362ye-evidence/callgrind-4b6e2c23bc`;
the final candidate rows are in
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362ye-evidence/callgrind-small-content-memo-v3`.
Both use `run_callgrind.sh`, one process per route, and the same ten route
arguments. The captures remain outside the repository.

| route | baseline Ir | candidate Ir | ratio |
|---|---:|---:|---:|
| `open` | 829,907,975 | 829,845,718 | 0.999925× |
| `first_slide` | 886,558,187 | 894,428,325 | 1.008877× |
| `repeat_slide` | 954,665,843 | 970,662,735 | 1.016757× |
| `first_text` | 899,838,108 | 907,901,737 | 1.008961× |
| `repeat_text` | 969,740,535 | 985,835,158 | 1.016597× |
| `write_text` | 877,897,998 | 882,989,662 | 1.005800× |
| `repeat_write_text` | 925,853,231 | 933,871,499 | 1.008660× |
| `unchanged_write` | 845,074,834 | 845,015,987 | 0.999930× |
| `append_commit_write` | 1,349,361,627 | 1,351,625,561 | 1.001678× |
| `replay` | 1,349,364,328 | 1,351,593,973 | 1.001652× |

The semantic-read instruction deltas are the bounded-memory path change, not
an additional repair scan. The e4a baseline reads its retained `content.xml`
string through `parse_*` and `write_slides_text_to`; the large-source current
owner uses the reader-backed `*_from_reader` variants and therefore reopens
and decompresses the member for each uncached query. The bounded writer memo
uses the existing validated `content_xml` snapshot only when the declared
member is at most 1 MiB, so repeated `write_text_to` calls reuse the decoded
member while larger sources remain streaming. This reduces the writer rows
from 1.012871×/1.024726× to 1.005800×/1.008660×. Its matched peaks are
694,399 bytes versus the e4a 693,231-byte baseline for both one and two
writes, with identical output lengths and hashes. The append lifecycle does
not have this tradeoff: its peak is 612,845 to 612,893 bytes (+48 bytes,
0.008%), with the same output and hash; no avoidable append scan or freshness
check was found.

Two broader memo designs were rejected. Applying the raw-content memo to
indexed slide reads raised `repeat_slide` to 1,582,211 bytes, above both the
e4a 1,500,683-byte peak and the reader-backed candidate's 990,985-byte peak.
Moving the ordinary text result cache to the first call violates the existing
source-read/freshness contract, which requires the second `text()` call to
stream before the thresholded cache hit. Neither change is included.

Native timing remains pending a quiet host window; these deterministic
instruction rows do not make a native-speed claim. A pinned-CPU diagnostic
ABBA was nevertheless captured for triage. It used 15 rounds, four timed
legs per route, and 30 samples per binary/route after pairing; the candidate
median divided by the baseline median was:

| route | candidate / baseline |
|---|---:|
| `open` | 0.9888× |
| `first_slide` | 0.9982× |
| `repeat_slide` | 1.0038× |
| `first_text` | 0.9974× |
| `repeat_text` | 1.0058× |
| `write_text` | 0.9900× |
| `repeat_write_text` | 0.9967× |
| `unchanged_write` | 0.9890× |
| `append_commit_write` | 0.9955× |
| `replay` | 0.9967× |

The run was pinned to CPU 5; host load was 24.85/22.44/23.91 at start and
22.27/22.08/23.72 at end. It is therefore diagnostic only, not quiet-host
acceptance. The raw 20,156-byte receipt is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/odp-0362ye-evidence/native-small-content-memo-v3-abba-cpu5-r15-correct.txt`
(SHA-256
`6f4bb23cb8c253e50dbc09671b0311a97063e5aaccba58cca6d8043193dca418`).
