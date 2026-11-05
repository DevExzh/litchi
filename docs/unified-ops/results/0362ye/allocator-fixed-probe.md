# 0362ye fixed admission probe

This is reduced allocator evidence for the fixed-buffer ODP admission probe.
The bounded writer-memo candidate is source commit `9834f169b1`. The generated
archive is the tracked `odp_0362ye_probe` fixture (2,000 slides, 4 MiB media).

| route | e4a baseline peak / allocated | current peak / allocated | output bytes | output SipHash |
|---|---:|---:|---:|---:|
| `open` | 693231 / 777196 | 83233 / 167925 | 0 | 15130871412783076140 |
| `first_slide` | 693231 / 3210784 | 89988 / 2681833 | 47 | 15128627771916252792 |
| `first_text` | 1500637 / 4672388 | 990939 / 4143437 | 105779 | 7807713202416953651 |
| `write_text` | 693231 / 2350086 | 694399 / 2432408 | 105778 | 18333823960124859280 |
| `repeat_write_text` | 693231 / 3922976 | 694399 / 4005676 | 211556 | 6976564693146433600 |
| `unchanged_write` | 6074765 / 6782817 | 5484747 / 6175554 | 4215428 | 12221861870545843138 |
| `append_commit_write` | 612845 / 5737255 | 612893 / 5821367 | 4215645 | 5546059324926701325 |
| `replay` | 612845 / 5737255 | 612893 / 5821383 | 4215645 | 5546059324926701325 |

The earlier fixed-buffer candidate binary is kept outside the repository at
`(deleted-coord-artifact)`
(3,599,992 bytes, SHA-256
`032047b9f559c291cc07b1052e5425b688a2e91784220c7ef57672d091136bf0`).
The final source build used for the current Callgrind and allocator rows is
`/home/zhuhe/code/litchi-worktrees/uo2-odp-heap-repair/target/release/examples/odp_0362ye_probe`
(3,621,528 bytes, SHA-256
`9e76c655cc0dab3d057c99a4bc8b5e5aa90a5ee3d53ffa4207a6ac4adada0a8d`).
The baseline binary is
`(deleted-coord-artifact)`
(3,096,104 bytes, SHA-256
`ecb034b0f57588c7d131ba9981d7004f5181ee214a88d14519f73660f35e0669`).

The candidate was built and run with:

```text
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
CARGO_TARGET_DIR=$CANDIDATE_TARGET TMPDIR=$ODP_TMP \
cargo build --offline -p litchi-odp --example odp_0362ye_probe --release
$CANDIDATE_TARGET/release/examples/odp_0362ye_probe open
```

The baseline used the same build command with its protected target and the
standalone source selected `route_before.rs` at compile time. The route
scripts accept an explicit fourth argument (`before` or `after`) and verify
that the binary's mode matches the requested route:

```text
docs/unified-ops/results/0362ye/run_allocator_routes.sh "$BIN" baseline "$OUT" before
docs/unified-ops/results/0362ye/run_callgrind.sh "$BIN" baseline "$OUT" before
```

The checked-in Cargo example is the current route by default, so an ordinary
all-target build and an unqualified example run need no environment. The
standalone record source retains `ODP_PROBE_ROUTE=.../route_before.rs` for the
historical baseline harness. This is allocation evidence; the route-local
Callgrind comparison below uses the same fixed baseline binary and the final
bounded-writer candidate.

The source hashes for this reproduction are:

```text
crates/litchi-odp/examples/odp_0362ye_probe.rs             3c681256eea27485ba71adff85c8c9090f81689ddb4a62675e65801a9614187d
docs/unified-ops/results/0362ye/odp_0362ye_probe.rs         9d709068b90f685fd45b367320296484020c4af45bd42da32d7ac13d5ac41650
docs/unified-ops/results/0362ye/route_before.rs             fc08a3d65ec38c9053b1927e25943762c06c6894026e186b011bfb0dab028bb9
docs/unified-ops/results/0362ye/route_after.rs              f27df012df0ca59cfc7845a659dd5e31c8a109ae1ebac8b07a8c888f17197dc6
docs/unified-ops/results/0362ye/run_allocator_routes.sh    91895e7b45b66ad9d2fd8f71994a9d353b4c27a5214781cb05491734d60ab4b1
docs/unified-ops/results/0362ye/run_callgrind.sh            475e0e8f18307b4a37483c7d28c02ff530560cf75c322871b9719ce1e5189c1c
```

## Route-local Callgrind

The fixed baseline rows are from
`(deleted-coord-artifact)`;
the final candidate rows are in
`(deleted-coord-artifact)`.
Both use `run_callgrind.sh`, one process per route, and the same ten route
arguments. The captures remain outside the repository.

| route | baseline Ir | candidate Ir | ratio |
|---|---:|---:|---:|
| `open` | 829,907,975 | 829,854,316 | 0.999935× |
| `first_slide` | 886,558,187 | 894,715,442 | 1.009201× |
| `repeat_slide` | 954,665,843 | 970,787,891 | 1.016888× |
| `first_text` | 899,838,108 | 908,198,425 | 1.009291× |
| `repeat_text` | 969,740,535 | 986,165,319 | 1.016937× |
| `write_text` | 877,897,998 | 883,060,231 | 1.005880× |
| `repeat_write_text` | 925,853,231 | 933,908,431 | 1.008700× |
| `unchanged_write` | 845,074,834 | 845,054,190 | 0.999976× |
| `append_commit_write` | 1,349,361,627 | 1,351,637,536 | 1.001687× |
| `replay` | 1,349,364,328 | 1,351,606,782 | 1.001662× |

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
`(deleted-coord-artifact)/native-small-content-memo-v3-abba-cpu5-r15-correct.txt`
(SHA-256
`6f4bb23cb8c253e50dbc09671b0311a97063e5aaccba58cca6d8043193dca418`).
