# XLSX retained-history paired adapter

The retained-history allocator route is `tools/perf-baseline/src/bin/xlsx_merge_history_alloc.rs`.
It is one source file for both revisions. The default build selects the W4
deferred owner; the `legacy-xlsx-source-history` feature selects the pre-W4
`workbook::source_merge` owner. Both branches use the same generated fixtures,
alternating merge/unmerge steps, replay and inverse byte checks, patch count,
and allocator snapshot interval.

The canonical source is recorded with SHA-256
`3929b1fbdaf4771d0667bc8a79cf0e274b72eebca6c064cb4a64d9c900c8f1b5` after the
final source edit and before the release build. The final edit makes the
legacy materialized inverse recipient mutable and removes an unnecessary
legacy-only mutability warning; both API variants compile from the same
adapter. Regenerate the temporary
baseline adapter from that source as follows; the archived baseline checkout
stays outside the repository:

```sh
cp tools/perf-baseline/src/bin/xlsx_merge_history_alloc.rs \
  /tmp/litchi-w4-baseline/tools/perf-baseline/src/bin/
python3 - /tmp/litchi-w4-baseline/tools/perf-baseline/Cargo.toml <<'PY'
from pathlib import Path
import sys

path = Path(sys.argv[1])
text = path.read_text()
text = text.replace(
    'xlsb-crud = ["litchi/xlsb"]\n',
    'xlsb-crud = ["litchi/xlsb"]\nlegacy-xlsx-source-history = []\n',
    1,
)
text = text.replace(
    '[[bin]]\nname = "litchi-perf-baseline-alloc"\n',
    '[[bin]]\nname = "xlsx_merge_history_alloc"\n'
    'path = "src/bin/xlsx_merge_history_alloc.rs"\n'
    'required-features = ["allocator-metrics", "legacy-xlsx-source-history"]\n\n'
    '[[bin]]\nname = "litchi-perf-baseline-alloc"\n',
    1,
)
path.write_text(text)
PY
sha256sum /tmp/litchi-w4-baseline/tools/perf-baseline/src/bin/xlsx_merge_history_alloc.rs
```

The baseline command enables both `allocator-metrics` and
`legacy-xlsx-source-history`; the candidate enables only `allocator-metrics`.
The measured counters are live bytes, high-water live bytes, allocation and
deallocation calls, reallocation calls, and allocated/deallocated bytes over
the interval that retains every patch. The W4 deferred publisher shares the
change record's `Arc` with the retained patch, so each retained commit has no
second closure allocation for its output endpoint. Exact replay and inverse
checks still stream through the recipient owner and its limits. For an
unobserved deferred after endpoint, the patch binds the recipient to the exact
before artifact and retains an immutable publisher whose catalog/order,
metadata, timestamps, save options and replacement bytes are all captured by
the change record. That deterministic publisher's single recipient-policy
stream is the authenticated after artifact; hashing the same archive in a
second prepass would duplicate the splice. Opened or independently observed
after endpoints retain their artifact fingerprint and use the exact restore
proof. Inverse replay authenticates its own before endpoint by the same rule.
The recipient source freshness check remains in the streaming write path, so a
source change during output is reported rather than silently accepted. RSS is
supplemental only.

The report has separate retained-history, verification, common replay, and
owner-admission allocator intervals. `verification_allocator` is the
complete caller workflow: every step opens an independent equal source-backed
recipient, replays forward and inverse, and serializes both outputs. Its W4
call chain is `SourceBackedWorkbook::from_read_at` →
`open_source_backed_package` → catalog and sheet-graph admission → deferred
replay/publication. The legacy adapter uses `OpcPackage::from_bytes` and its
materialized patch API, so that interval is retained as an unmatched workflow
result; it does not construct the new read-only facade. The report also records
`source_owner_admission_allocator`, which repeats source owner construction
and selected worksheet admission inside the measured interval for each
revision, and `common_replay_allocator`, which constructs both recipients
before the interval and measures the shared forward/inverse replay and
publication work. The latter keeps the exact replay and inverse byte oracles
while removing the unequal owner-admission cost from the common route.
`common_replay_apply_forward_allocator` and
`common_replay_apply_inverse_allocator` split the same recipient-only apply
work before serialization, so deferred authentication allocations in inverse
replay are attributable separately from publication and output buffering.

An earlier probe was rejected: it independently opened a source owner for each
retained step and materialized every commit output while measuring history.
Its candidate live and peak bytes were therefore contaminated by the probe's
per-step owners. The final adapter reuses the two fixture owners for commit
history and retains only the patch records; its replay oracles remain outside
the retained commit output. The rejected JSON remains outside the repository
at `/tmp/litchi-w4-history-abba-20261021.json` for audit continuity.


## Reproducible pre-W4 allocator pair

The historical frozen baseline remains diagnostic because its extracted source
archive has no complete revision provenance. A reproducible pre-W4 baseline is
now prepared from the exact current-trunk integration commit
`e63a1fe3534bcaaae2ce367339f7f8275877baef` in the external worktree
`/home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-alloc-v1`. A file-filtered
base archive is retained outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-prew4-allocator-base-v2.tar`
(206,110,720 bytes, SHA-256
`c9281495a28f5abaea8326091071c2f4e77a5015a38596709c287fcd74cee42a`); its
selected path list is next to it with SHA-256
`d69e80ecb137da9d9f040afa0fe26e3f2b46a734279b758b5942e81ce0469d27`. The
archive contains the exact base commit's Cargo manifest, `.cargo`, crates,
perf harness, and build lint/toolchain manifests, and excludes documentation,
media, test data, unrelated tools, and all performance captures. The only source
overlay is this adapter at SHA-256
`3929b1fbdaf4771d0667bc8a79cf0e274b72eebca6c064cb4a64d9c900c8f1b5`; the
Cargo manifest adds the `legacy-xlsx-source-history` feature and registers the
adapter binary. The release build uses the W4 environment shim,
`--manifest-path tools/perf-baseline/Cargo.toml`, binary
`xlsx_merge_history_alloc`, and features
`allocator-metrics,legacy-xlsx-source-history`. The final candidate uses the
W4 worktree, the same adapter, and `allocator-metrics`.

Protected release binaries are outside the repository:

- pre-W4 base: `xlsx_merge_history_alloc-base-e63a1fe353-v1`, SHA-256
  `12ff3a682f35f5c2292ed9232afe59aa78f3393c548584ab863d45e7802b8014`;
- W4 candidate: `xlsx_merge_history_alloc-candidate-fee2cdf-v2`, SHA-256
  `7201da75e2a91abb80e5fa9727c488d00349da254e4c95f61e693305dd8df539`.

Both builds exited successfully. The allocator comparison below is a byte and
call-count attribution run, not a wall-time measurement; native timing still
requires a separate quiet grant.

The 2a-to-fee common replay allocation difference is localized rather than
hidden in the aggregate. Forward apply allocation fell from 1,009,920 to
51,200 bytes, while inverse apply rose from 1,359,872 to 55,034,496 bytes.
Publication-forward allocation fell from 55,652,224 to 54,911,872 bytes, and
publication-inverse stayed at 319,488 bytes. Thus the 2a-to-fee route's lower
call count and lower retained peak did not justify accepting its inverse
allocation path by itself. The reproducible pre-W4 comparison below resolves
the aggregate attribution; the inverse apply cost remains separately reported.

## Reproducible pre-W4 source-family and compact pair

The same provenance discipline was applied to the complete source-family
orchestrator. The base is the exact pre-W4 commit
`e63a1fe3534bcaaae2ce367339f7f8275877baef`; the final candidate is this
worktree at `6800ddc682f29bbcf1f56f83b2725728241aeae`. The base perf lock has
SHA-256
`0e1810824db584d171c16b45930776812a6e44c36e62c3e1e5b52e49a86da1cd`; the
candidate lock has SHA-256
`5f3e7d34cd8b07fb12d10f29fc41c2c42146907fd2b0e6bc8528b9eb772fc79c`.
The lock difference is the expected revision graph: the pre-W4 lock retains
`litchi-core` on `litchi-odraw` and `litchi-biff`/`litchi-ograph` on `litchi-ops`,
while the candidate removes those edges and adds `sha2 0.11` to the
`litchi-xlsx` dependency graph. Neither main child selects a Cargo feature beyond
the features declared by the manifest dependencies.

The e63 checkout has the six-line temporary manifest overlay used by the
shared retained-history adapter. Its working `tools/perf-baseline/Cargo.toml`
has SHA-256
`d89eff26ca16ebb28e4b153ef790eed57379a0ad1516775dd8046de2e1ecc545`, and the
overlay diff has SHA-256
`8bed42293a1145a59b4570688ae984fdc635f4858ca6ec0308bfcabad2b`. The adapter
source remains `3929b1fbdaf4771d0667bc8a79cf0e274b72eebca6c064cb4a64d9c900c8f1b5`.
That adapter is not compiled into either main source-family child; the overlay
is retained so the archived checkout and all related build recipes share one
manifest. The candidate manifest has SHA-256
`d3146586bf1325679a73202d9dafa61573a492cc2750668b5b408b7ab010febe`.
The e63 perf harness already contains all six eager/source case names used by
the final driver, so no synthetic source-family compatibility adapter or old
binary was reused; only the retained-history allocator adapter uses the
temporary overlay.

The exact release commands were, with the W4 environment shim and distinct
target directories:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-target-main-v1
cargo build --release --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline
cargo build --release --manifest-path tools/perf-baseline/Cargo.toml \
  --bin xlsx_merge_source_abba

source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-w4-alloc-target-v2
cargo build --release --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline
cargo build --release --manifest-path tools/perf-baseline/Cargo.toml \
  --bin xlsx_source_family_abba --bin xlsx_merge_source_abba
```

The protected binary identities are:

- e63 main child: `e0245be90b20afc18c1a13554fc92a1fdde27a7e0e6ee9c781bcfcbba5f78cbf`;
- final main child: `b2eab1f7fe0a281c3b040c4404d4071ba3e45181d9dbd0def4e7987f5f4811ba`;
- e63 merge child: `9d746102c663324b06b93530de3e93094c6af303993eab83e5328183fa9ec176`;
- final merge child: `41604f16ec4516f06a66b4ab9a4ccd002d621a55ed7cf0f2783ab41b47067d0d`;
- final source-family driver: `bc61708b336e36138aa459bed8ef170f4beb6e21427436ea5c3377896df142e4`.

The driver source is SHA-256
`165de1ef06c37811f49e6ec66bb85a42a72ff33cac9b35b80646922589e1b7e0`; its
merge-child source is
`604805876dae86480eef0a0ea7a90cecd588bff2452bfe354ce9a8fd629cea71`.
Preflight with the final driver, e63 base children, and final children
completed with 24 family cases and four compact children. Its receipt is
outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-source-family-preflight-head-v2/preflight.json`.

The e63 and final allocator reports use the same fixture hashes and exact
128/128 forward and inverse checks. In the comparable common replay interval,
allocated bytes changed from 115,486,976 (e63) to 110,317,056 (final),
allocation calls from 107,776 to 23,552, and post-interval peak live bytes
from 866,252 to 672,471. The aggregate therefore improves by 4.48%; the
earlier 2a-to-fee +89% figure was not a real regression of this workload.
It came from comparing the W4 adapter's two different deferred implementations
under a mismatched common route. The phase split explains the discrepancy:
apply bytes rise from 6,493,952 to 55,085,696, almost entirely in inverse
authentication, while publication bytes fall from 109,259,648 to 55,231,360,
almost entirely because inverse publication no longer rematerializes the
artifact. The total common route is lower and exact, but the inverse apply
allocation remains a separately reported cost for any later wall-time review.

The complete allocation JSON pair is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-source-allocation-v1/`.
The base and candidate JSON SHA-256 values are
`9fbaff2f43b8115834e7c12bd8be33306a1f2f33d9ada5bb34145b02296ee0f0` and
`083fa101f3a8a88360fc8114a28046090ce1528c3388fcbc15a23472371ce018`,
respectively.
No native timing is claimed from this attribution run.

The prepared native protocol is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-source-family-abba-v1.commands`.
Its SHA-256 is
`e2b0b79e365e311e3eec966febcb9273080584f2d72cc0d9338fc688a4b6bd51`. It uses
only the protected binaries above and the final source-family driver.
The focused timing legs select `page_margin` and `row_visibility`, use two
warmups and 64 measured child samples, and retain every ABBA leg and route
checkpoint. The batch size is chosen to put the source and medium-row legs
above 150 ms of measured child work; the receipt must report each leg's actual
duration before any performance claim. A separate page-margin run enables
`--compact` with 20 observations for compact forward/inverse, stale, foreign,
and exact-overlay checks. The full six-family `--samples 64 --compact` command
is retained in the same runbook for a later grant when the coordinator wants
the complete matrix. No native timing has been started.

## Batched source-family harness correction

The first native attempt with the Rust xlsx_source_family_abba driver is
unaccepted. That driver invoked the protected child once for every leg with
--samples 1 and retained a route checkpoint only after both implementations
finished. The coordinator stopped it after a complete partial ABBA block when
the missing batching and checkpointing were found. Its temporary child JSON
files were deleted by TemporaryJson; no elapsed observations survived. The
external diagnostic receipt records the protected binary identities and this
loss at
/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-source-family-abba-v1/attempt-aborted-20261021.json (SHA-256
a521dda6ae331e3db55a530baa26e7efea5502cf0ad3038194c89f70e95fc25a). It is
not a timing result.

The replacement adapter is tracked at
tools/xlsx_source_family_abba_batched.py. It invokes each protected child
with its existing --samples N complete-operation batch, independently
normalizes N for each endpoint and revision, and uses five ABBA blocks by
default. Its target is 500 ms of summed measured operation time per leg, with
150 ms as the hard minimum. Two explicit warmup child calls per endpoint are
recorded separately and excluded from measured sums. Every calibration,
warmup, measured, short, invalid, or failed child result is appended to an
fsync'ed events.ndjson journal before the next child starts; each completed
ABBA block has an atomic checkpoint. Raw child reports, output hashes, sink
bytes and largest-write bounds, source read counters, per-operation samples,
exact N, and per-leg sums are retained. Compact merge/unmerge children are
repeated to the same summed-span floor and report per-leg primary-time sums;
those repeated spans are not presented as one contiguous child clock.

The adapter has a --plan mode that invokes no child; `--preflight-only` is not
an alias for it. A timing run uses only
protected copies outside the repository, for example:

~~~sh
python3 tools/xlsx_source_family_abba_batched.py \\
  --before-bin /tmp/litchi-w4-final-binaries-20261021/litchi-perf-baseline-base-e63a1fe353-v1 \\
  --after-bin /tmp/litchi-w4-final-binaries-20261021/litchi-perf-baseline-candidate-6800ddc682-v1 \\
  --merge-before-bin /tmp/litchi-w4-final-binaries-20261021/xlsx_merge_source_abba-base-e63a1fe353-v1 \\
  --merge-after-bin /tmp/litchi-w4-final-binaries-20261021/xlsx_merge_source_abba-head-6800ddc682-v1 \\
  --compact --blocks 5 --warmup 2 --target-ms 500 --minimum-ms 150 \\
  --out-dir /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-source-family-abba-batched-v1
~~~

The adapter refuses an existing nonempty result directory, records its own
runner hash, every protected family/compact binary identity, CPU affinity, and
durable compiler/profiler process scans before the run and between ABBA blocks.
A busy scan aborts before another child starts. Child `elapsed_ns.samples`
already contains one value per complete operation: the values are retained
unchanged, each leg reports `sum_elapsed_ns` and `mean_elapsed_ns`, and each
block reports the paired mean ratio after/before. Source read vectors and sink
geometry are checked for their requested length and stability within each
revision; before and after signatures remain separate, while output hashes are
required to match across the paired route. The focused adapter tests cover
different endpoint sample counts, real-shaped source vectors, one-time mean
normalization, per-revision stability, fresh directories, and rejection of the
removed preflight alias.

The corrected adapter source SHA-256 is
96d3d7a34da55fe3abfde4e502d5d23ff20c5d7e775dbb85d82df10a949b3ac4. It has
passed `python3 -m py_compile` and the focused fake-report tests. No native
timing has been run with the corrected adapter.
