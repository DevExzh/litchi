# 0362y — Keep the XLS eager-cell walk concrete

The legacy XLS adapter no longer stores its eager worksheet cell walk as a
private `Box<dyn Iterator>`. `litchi_xls::worksheet::StoredCells<'a>` names
the existing zero-copy `BTreeMap` values iterator, and
`format::xls::{owner,children}` carry that concrete iterator through the
cursor state. The worksheet's values, order, and lifetime are unchanged; no
cell data is collected or copied.

This closes the memo 29 residual without changing the source-backed cursor,
grid fallback, path behavior, or error classification. The source-backed XLS
commit path now retains its validated overlay plan and defers opening the
composed workbook until the first semantic read. Write-only publication still
streams the checked plan directly, while the first read opens the target once
and caches it for subsequent projections. Commit-time source, span,
fingerprint, and numeric readback validation remains unchanged. The
coordinator should qualify write-only, first-read, and read-after-write costs
against the former forward-only plan before claiming the F3 saving.

The deferred cell family fences the original dynamic source before and after
planning and before and after each deferred target materialization/publication;
staging's owned cell snapshot is not a freshness oracle. A stale
stage-to-commit therefore returns the retained family `Io` source-change
reason, while a stale deferred `write_to`/`save` retains the typed
`OutputError::Package(Error::Source(SourceChanged { .. }))` boundary. The
non-deferred source-version path keeps its existing fast path.

Validation was limited to formatting and static diff/artifact guards in this
leaf. The coordinator should run the XLS and unified-ops checks on the merged
assembly.

## F3 qualification lane (executable, not measured)

The fair implementation comparison is the exact source snapshot at
`9b1f96d6dce1506dc9d0080f0696b42ba592c397` (the concrete eager-cell iterator
is already present, but the deferred source-backed cell reopen is not) against
an explicitly pinned candidate source ref. The default candidate ref is
`0b8f68d7e9d3c97cfa62cedceacdeb2ad5f4ba76`, which contains the two production
freshness fixes `8d491b1e8f` and `ef6e6ea253` plus the typed deferred-cell
refusal, protected-source, and baseline-disposition fixes `157eb7c8d2`,
`c7ed9d3abd`, and `a443967991`. The candidate ref may be replaced on the
command line with the final assembled commit, but the receipt must retain that
exact ref and the required production commit ancestry. The historical baseline
remains exactly `9b1f...`; production files are never cherry-picked into it. No
result may be attributed solely to F3 without the full candidate source identity
and production control. The harness source
`tools/perf-baseline/src/xls_numeric.rs` is taken from the frozen common
harness ref for both binaries.

The existing six opt-in numeric selectors remain the source of the corpus and
fidelity oracles:

```text
xls_numeric_eager_number_edit_save
xls_numeric_source_backed_number_edit_save
xls_numeric_eager_rk_mulrk_edit_save
xls_numeric_source_backed_rk_mulrk_edit_save
xls_numeric_plan_only_number_edit_save
xls_numeric_plan_only_rk_mulrk_edit_save
```

The executable three-route probe was introduced by `d8fe3bcacb`; its
freshness/refusal/replay gates plus the first ABBA runner are committed as
`b03aaf1aac` in
`tools/perf-baseline/examples/xls_f3_qualification.rs`, with its library
implementation in `tools/perf-baseline/src/xls_f3_qualification.rs`. The
shared `tools/perf-baseline/src/xls_f3_corpus.rs` module owns both canonical
fixture builders. The isolated `tools/perf-baseline-f3` manifest reuses that
module and the existing probe source, so it does not compile unrelated DOCX or
PPTX baseline modules or duplicate workload construction.

The common harness is frozen at
`0b8f68d7e9d3c97cfa62cedceacdeb2ad5f4ba76`. Prepare both source snapshots at
their exact production refs, then overlay only the allowlisted probe files
from that frozen commit. This is a content overlay, not a cherry-pick: it
keeps the baseline production tree at `9b1f...` and keeps the candidate
production tree at its explicitly selected ref while making the binaries use
one identical harness. The overlay is deterministic and auditable:

```sh
REPO=/home/zhuhe/code/litchi-worktrees/uo2-xls-f3-qualification
HARNESS_REF=0b8f68d7e9d3c97cfa62cedceacdeb2ad5f4ba76
BASE=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/source-base-9b1f96d6dc
CAND=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/source-candidate-0b8f68d7e9
PROBE_PATHS=(
  tools/perf-baseline/src/xls_f3_qualification.rs
  tools/perf-baseline/src/xls_f3_corpus.rs
  tools/perf-baseline/examples/xls_f3_qualification.rs
  tools/perf-baseline/src/xls_numeric.rs
  tools/perf-baseline/src/lib.rs
  tools/perf-baseline-f3/Cargo.toml
  tools/perf-baseline-f3/Cargo.lock
  tools/perf-baseline-f3/src/lib.rs
  tools/perf-baseline-f3/examples/xls_f3_qualification.rs
)
for SOURCE in "$BASE" "$CAND"; do
  for PATH_IN_SOURCE in "${PROBE_PATHS[@]}"; do
    mkdir -p "$SOURCE/$(dirname "$PATH_IN_SOURCE")"
    git -C "$REPO" show "$HARNESS_REF:$PATH_IN_SOURCE" > "$SOURCE/$PATH_IN_SOURCE"
  done
done
```

The runner checks the exact frozen blob hashes for every overlay path, allows
only those paths to be dirty or newly added, compares the production manifest
with the selected source ref, and compares the resulting production-plus-
overlay source manifest with its expected identity before launching a child.
The candidate must descend from all five pinned production commits:
`8d491b1e8f084effde712651e07d460ba044cb1f`,
`ef6e6ea25355e56571f870a4cefd1675fdb60366`,
`157eb7c8d250568acb9dbc83bbeb0f9e675d1ea5`, and
`c7ed9d3abdc3f9e3ee152c12129f194d82b66042`, and
`a443967991dfbdf7529c198abf2f6d24211e87ac`; the control source ref is fixed to
`9b1f96d6dce1506dc9d0080f0696b42ba592c397`. Before and after every
child it records the selected CPU's `/proc/stat` counters, load average,
runnable count, and selected-CPU runnable processes. A scan is quiet only when
no heavy process is present, one-minute load is at most 1.0, runnable count is
at most one, and no other process is running or uninterruptibly sleeping on
the selected CPU.

The provenance guard has a read-only real-checkout test and temporary Git
fixtures for both an admitted exact overlay and rejected production/unlisted
path changes:

```sh
python3 -m unittest -v tools/test_xls_f3_qualification_abba.py
```

The Number corpus edits `Untouched!E21`, `Number 42` to `43`; the native
RK/MulRK corpus edits one RK and two MulRK values with fixed-width integer
replacements. Every fresh sample opens a new owner and stages the same edit.
The existing source-backed selector is the **write-only** route: it commits
and immediately calls `commit.package().write_to(&mut sink)` without a semantic
read of the committed owner. It already records separate edit, set, commit,
publication and total vectors. The plan-only selector is not a first-read
control; it is another write-only publication contract.

The executable F3 probe accepts an implementation disposition, one required
route and an optional corpus: `--implementation baseline|candidate`,
`--route write-only|first-read|read-after-write` and
`--corpus number|rk-mulrk|all` (default `all`). It emits one CSV header followed
by rows with `route,corpus,sample,nanoseconds,edit_ns,set_ns,commit_ns,write_ns,read_ns,output_bytes,output_sha256,batch`.
The pinned `9b1f...` baseline has three exact historical refusal dispositions:
the signed `edit_cells` object-layer `InvalidFormat`, the structural
source-backed plan's `UnsupportedFeature` text, and the RK/MulRK
`set_number` wrong-family `UnsupportedFeature` text for the selected fixture
coordinate. It also records five pinned baseline freshness dispositions that
predate the candidate fences: stage→mutate→commit and same-owner forward and
inverse replay return a committed owner with nonempty publication, except that
the inverse replay's publication returns exactly
`OutputError::Package(Error::Source(SourceChanged { .. }))`; a
mutate→committed-first-read returns the replacement value; and
mutate→committed-write returns nonempty publication. The
candidate must return the corresponding typed refusals. The runner passes the
disposition explicitly to every child and records this allowlist in the
receipt; no arbitrary error or broad `is_err()` path is accepted.
The probe has already added three explicit route labels over this same
preparation:

| route | timed sequence | route-specific oracle |
| --- | --- | --- |
| `write-only` | edit, stage, commit, `write_to` | no semantic call on the committed owner before publication |
| `first-read` | edit, stage, commit, one `cell_value_by_index` on the committed owner | the first read returns the replacement and a second untimed read returns the same value, proving the materialized owner is cached |
| `read-after-write` | edit, stage, commit, `write_to`, then one semantic read on that same committed owner | the output is complete before the first read, and the read still returns the replacement |

The first-read route must write and verify its output outside the timed
sequence; the read-after-write route includes both write and first read in the
timed sequence. This keeps the write-only result honest while measuring where
the candidate moves the target reopen. A fresh source owner and a fresh
committed owner are required for every sample; no route may reuse a
materialized owner across samples. The source-backed baseline already opens
the composed target during commit, so its first-read cost is expected to be
paid in the commit phase; the candidate pays it at the first semantic read.

All three routes retain the existing untimed gates: exact output digest and
bounded sink byte/write-call counts; untouched CFB streams and topology;
selected value and storage-family readback through `cell_values`; raw
`litchi_xls::raw::Reader` readback; forward replay and inverse replay to the
original bytes; refusal when applying a patch to its committed owner; partial
sink failure; exact no-op and source/target fingerprint behavior; protected,
signed, macro, encrypted and unsupported-value refusals matched to their exact
`RefusalKind`, operation and fixed reason; the exact wrong-storage-family
error; an unchanged mutable-source control; stale read, stage-then-mutate
commit, and mutate-after-commit write `SourceChanged` oracles; source bytes
unchanged from the original fixture; typed stale-recipient `PatchConflict`;
typed `IncompleteOutput` with an authored prefix; and the real-producer
`54016.xls` RK/MulRK forward and inverse gate. Any route that omits one of
these checks is diagnostic only and cannot qualify F3.

The qualification receipt is expected to make this checklist auditable:

| review item | executable evidence |
| --- | --- |
| typed protected, macro, unsupported and wrong-family gates | probe refusal `kind`/operation/reason or exact fixed-family error |
| freshness controls | unchanged read, stale read, stage→mutate→commit, mutate→committed-first-read, and mutate→committed-write cases |
| source preservation | source owner writes equal the original fixture before and after each sample |
| provenance and identities | exact source refs, frozen harness ref, overlay-only status, expected production manifest, expected production-plus-overlay source manifest, and exact probe-file hashes |
| matrix coverage | six unique route/corpus cells × four unique A1/B1/B2/A2 legs; subsets require diagnostic mode |
| timing floor | independent control/candidate calibration and every retained leg at least 250 ms |
| resource isolation | observed child affinity/RLIMIT_AS and `/proc` load/process scans before and after every child |
| CSV protocol | exact header and exact field set on every row, with stable sample/output identities |

The old external W5 probe at
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-final-20261021`
is a method template only: its `xls_cells` and `xls_cells_stage` routes use
`litchi_xls::Package` and are eager. It must not be used as the deferred F3
binary. Keep all new source snapshots, binaries, captures and reports outside
the checkout, for example:

```text
/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/
  source-base-9b1f96d6dc/
  source-candidate-0b8f68d7e9/
  target-base-release/
  target-candidate-release/
  binaries/
  abba/
  heap/
```

Build each source snapshot with a distinct target and one worker after the
coordinator grants the slot (the command is recorded here, but was not run in
this leaf):

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/env.sh
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/target-base-release \
  cargo build --offline --locked --release \
  --manifest-path /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/source-base-9b1f96d6dc/tools/perf-baseline-f3/Cargo.toml \
  --example xls_f3_qualification
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/target-candidate-release \
  cargo build --offline --locked --release \
  --manifest-path /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/source-candidate-0b8f68d7e9/tools/perf-baseline-f3/Cargo.toml \
  --example xls_f3_qualification
mkdir -p /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/base \
  /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/candidate
cp /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/target-base-release/examples/xls_f3_qualification \
  /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/base/xls_f3_qualification
cp /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/target-candidate-release/examples/xls_f3_qualification \
  /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/candidate/xls_f3_qualification
sha256sum \
  /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/base/xls_f3_qualification \
  /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/candidate/xls_f3_qualification
```

The F3 probe binary name above is the implemented qualification entry point.
The existing numeric selectors remain the retained corpus/oracle control and
are run separately with the README command. No evidence belongs under
`docs/performance/results/` until a reduced receipt is ready.

The native run is four process legs per route and corpus in strict A1/B1/B2/A2
order, with baseline A and candidate B using the same route, sink and
affinity. The runner calibrates both endpoints independently, then uses the
larger admitted batch for both sides; it rejects every leg shorter than 250 ms,
retains every raw leg, and records process/load scans before and after each
child. The executable runner is `tools/xls_f3_qualification_abba.py`; it fails
closed on source-ref mismatch, changes outside the exact overlay allowlist,
frozen probe-hash mismatch, production-manifest mismatch, source-manifest
mismatch, missing required production ancestry, child errors, malformed or
extra CSV fields, unstable output digests, equal control/candidate binary
hashes, and any order other than A1/B1/B2/A2. It writes each child
stdout/stderr and a JSON receipt containing source refs, frozen harness ref,
overlay and production/source manifests, binary, runner-script, host,
calibration, affinity and memory-scope hashes. A coordinator-granted quiet CPU
is represented below as `${CPU}` and must be fixed in the manifest. A subset
requires `--diagnostic` and its receipt has `qualification_eligible: false`;
only the default six-cell matrix is eligible:

```sh
BASE=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/base/xls_f3_qualification
CAND=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/candidate/xls_f3_qualification
OUT=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/abba
python3 /home/zhuhe/code/litchi-worktrees/uo2-xls-f3-qualification/tools/xls_f3_qualification_abba.py \
  --control-binary "$BASE" \
  --candidate-binary "$CAND" \
  --control-source /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/source-base-9b1f96d6dc \
  --candidate-source /home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/source-candidate-0b8f68d7e9 \
  --output-dir "$OUT" \
  --cpu "${CPU}" \
  --memory-limit-bytes 2147483648 \
  --warmup 20 --samples 40 \
  --harness-ref 0b8f68d7e9d3c97cfa62cedceacdeb2ad5f4ba76 \
  --control-source-ref 9b1f96d6dce1506dc9d0080f0696b42ba592c397 \
  --candidate-source-ref 0b8f68d7e9d3c97cfa62cedceacdeb2ad5f4ba76 \
  --candidate-production-commits 8d491b1e8f084effde712651e07d460ba044cb1f,ef6e6ea25355e56571f870a4cefd1675fdb60366,157eb7c8d250568acb9dbc83bbeb0f9e675d1ea5,c7ed9d3abdc3f9e3ee152c12129f194d82b66042,a443967991dfbdf7529c198abf2f6d24211e87ac
```

The runner emits A1/B1/B2/A2 for each route/corpus pair, and stores its script
hash, exact frozen overlay-file hashes, production and combined source
manifests, binary hashes, source refs, CPU, memory scope, calibration attempts,
quiet scans, raw stdout/stderr and output digest in
`$OUT/xls-f3-abba-receipt.json`. The admitted common batch is measured by the
runner itself, with both endpoint calibration legs required to reach 250 ms.
It must not report a speedup from the old eager `ole_w5_probe`.

For resource evidence, use the same release probe binaries and run direct
heaptrack children (not a wrapper hidden behind `taskset`) in fresh
directories. The probe intentionally has no counting-allocator mode in this
source-only preparation; heaptrack is whole-process evidence with its own
instrumentation overhead, not a native route-timing result:

```sh
HEAPTRACK=$(command -v heaptrack)
HEAPTRACK_PRINT=$(command -v heaptrack_print)
for side in base candidate; do
  binary=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/binaries/$side/xls_f3_qualification
  implementation=baseline
  if [ "$side" = candidate ]; then implementation=candidate; fi
  for route in write-only first-read read-after-write; do
    for corpus in number rk-mulrk; do
      dir=/home/zhuhe/code/litchi-worktrees/.coord/xls-f3-qualification/heap/$side/$route/$corpus
      mkdir -p "$dir"
      "$HEAPTRACK" --record-only --output "$dir/heaptrack" "$binary" \
        --implementation "$implementation" --route "$route" --corpus "$corpus" \
        --warmup 0 --samples 1 --batch 1 \
        > "$dir/probe.stdout" 2> "$dir/heaptrack.stderr"
      capture=$(find "$dir" -maxdepth 1 -type f \( -name 'heaptrack.*.zst' -o -name 'heaptrack.*.gz' \) -print)
      "$HEAPTRACK_PRINT" -f "$capture" -l -H "$dir/histogram.tsv" \
        > "$dir/heaptrack-print.txt" 2> "$dir/heaptrack-print.stderr"
    done
  done
done
```

Heaptrack output is whole-process allocation evidence with instrumentation
overhead, not native timing. The receipt must include allocation calls,
histogram bytes, peak heap/RSS, retained bytes, route output hash, binary and
probe hashes, and all failed or discarded captures. These artifacts remain in
the external qualification directory; none is committed here.
