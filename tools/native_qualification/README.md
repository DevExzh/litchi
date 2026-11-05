# Safe W4b/W5 native qualification runner

`tools/native_qualification_runner.py` replaces the historical W4b and W5
shell launchers for final execution.  It does not contain a source commit or
binary identity.  The caller supplies the final baseline/candidate worktrees
through the manifest environment, so a returned source repair cannot silently
reuse an old receipt.

The two manifests preserve the recorded workloads and order:

- `manifests/w4b-xlsb.json`: `xlsb_w4b_probe`, `sheet_dispatch_bench`, then
  `xlsb_crud` for each leg, 31 rounds, and
  `before after after before`.
- `manifests/legacy-w5.json`: `legacy_w5_source_probe`, then
  `legacy_w5_probe` for each leg, 31 rounds, `LITCHI_PROBE_BATCH_ROUNDS=7`,
  and the same ABBA order.

The legacy format preparation manifests reuse this same guarded runner:

- `manifests/legacy-rtf.json` binds the existing RTF refusal probe to five
  guarded ABBA blocks and the eight-route validator;
- `manifests/legacy-ole-properties.json` binds the existing OLE property
  probe's `read` and `edit` routes with 20 warmups and 500 retained samples;
- `manifests/legacy-xls-properties.json` binds the existing
  `ole_metadata_bench` executable and its tracked DOC/XLS/PPT fixtures to the
  derived pre-lazy XLS source parent.

These manifests require full source/candidate paths, locks, and candidate SHAs
from the coordinator.  Their route/output, identity, quiet-host and cap
checks are implemented by `legacy_native_validator.py`; the validator keeps
allocator instrumentation as a separate non-claimable phase.  The complete
source-fence and future execution recipe is recorded in
`docs/unified-ops/0362ul8-legacy-native-protocol.md`.

Before a final build, prepare the historical source consumers in external
worktrees.  The accepted W5 baseline is `7afba5f146aa1a16eba73ebed2144dd041b33fb5`;
the accepted W4b baseline is the synthetic merge tree
`684a028b2ae30a3a110b87bf6743fc598c6691a2`, reproduced by merging the
recorded W4-XLSX ref `635b884c7c5372a4aa05ea3f13c77eccc0469598` with the
recorded trunk ref `f2bbbc8d07ab5da26c180f1ff74813d39f632d54`.  The historical
measured heads are W5 `e3e805f2dd26db78b8e951fa920275760e87a22b` and W4b
`1af8b42cb6c4a12c3a048cc36b6342ef07498d1d`.  The current clean source refs
from the individual wave closures (`83bfc60b42ad7b6830c81bbc30e1238efadef907`
for W5 and `b2e21c0ab481ed7ff8e3ce07e50fcb3c9572fc90` for W4b) are historical
provenance only, not final candidates.  The coordinator must select repaired,
assembled candidate commits after the final merge; the preparer requires each
selected candidate's full commit SHA and records it in its provenance.

The baseline trees intentionally predate their wave probe consumers.  Create
the W4b synthetic worktree, create a detached W5 worktree at the accepted
tip, then run this source-only preparer.  It copies the W4b probe, applies the
recorded one-line `Cell` import to `xlsb_crud`, and generates the W5 probe with
the recorded `routes_before.rs`; it writes hashes and source identities to an
external provenance JSON and runs no build command:

```sh
W4B_TREE=$(git merge-tree --write-tree \
  635b884c7c5372a4aa05ea3f13c77eccc0469598 \
  f2bbbc8d07ab5da26c180f1ff74813d39f632d54)
test "$W4B_TREE" = 684a028b2ae30a3a110b87bf6743fc598c6691a2
W4B_COMMIT=$(printf '%s\n' 'record-0360b synthetic baseline' | \
  git commit-tree "$W4B_TREE" -p 635b884c7c5372a4aa05ea3f13c77eccc0469598 \
  -p f2bbbc8d07ab5da26c180f1ff74813d39f632d54)
git worktree add --detach /external/sources/w4b-before "$W4B_COMMIT"
: "${W4B_AFTER_REF:?set to the coordinator-selected full repaired assembled W4b candidate SHA}"
: "${W5_AFTER_REF:?set to the coordinator-selected full repaired assembled W5 candidate SHA}"
git worktree add --detach /external/sources/w4b-after "$W4B_AFTER_REF"
git worktree add --detach /external/sources/w5-before \
  7afba5f146aa1a16eba73ebed2144dd041b33fb5
git worktree add --detach /external/sources/w5-after "$W5_AFTER_REF"

python3 tools/native_qualification/prepare_w4b_w5_inputs.py \
  --w4b-before /external/sources/w4b-before \
  --w4b-after /external/sources/w4b-after \
  --w4b-after-ref "$W4B_AFTER_REF" \
  --w5-before /external/sources/w5-before \
  --w5-after /external/sources/w5-after \
  --w5-after-ref "$W5_AFTER_REF" \
  --output /external/native-final/source-preparation.json
```

Use the analyzer scripts already tracked by the selected candidate sources:
`$W4B_AFTER_SOURCE/docs/unified-ops/results/0360b/analyze.py` and
`$W5_AFTER_SOURCE/docs/unified-ops/results/0359a/analyze.py`.  The fixture is
the tracked `$W4B_AFTER_SOURCE/test-data/poi/test-data/spreadsheet/testVarious.xlsb`.

Use a new external output directory for each attempt.  The target directories
must also be external to both source worktrees.  The runner refuses an output
directory containing an existing phase journal/receipt and never deletes an
existing target, binary, log, or measurement output.

Example W4b setup (replace every path with the final frozen paths):

```sh
export NATIVE_QUAL_OUTPUT=/external/native-final/w4b-20261023-run1
export W4B_BEFORE_SOURCE=/external/sources/w4b-before
export W4B_AFTER_SOURCE=/external/sources/w4b-after
export W4B_TARGET_ROOT=/external/native-final/targets/w4b-run1
export W4B_XLSB_FIXTURE=/external/fixtures/testVarious.xlsb
export LITCHI_COORD_BIN=/home/zhuhe/code/litchi-worktrees/.coord/bin
export LITCHI_NATIVE_QUAL_ROOT="$(git rev-parse --show-toplevel)"
export W4B_ANALYZER="$W4B_AFTER_SOURCE/docs/unified-ops/results/0360b/analyze.py"

python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4b-xlsb.json \
  --phase build
python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4b-xlsb.json \
  --phase abba --cpu "$CPU"
```

Use the analogous `W5_BEFORE_SOURCE`, `W5_AFTER_SOURCE`, and
`W5_TARGET_ROOT` variables for `legacy-w5.json`, and set
`W5_ANALYZER="$W5_AFTER_SOURCE/docs/unified-ops/results/0359a/analyze.py"`.
The manifest runs the tracked adapter from the
final runner checkout (`LITCHI_NATIVE_QUAL_ROOT`), not from either historical
source tree: it executes that analyzer, parses its JSON stdout against the W5
analysis schema, and writes the unchanged analysis under the qualified
`summary.json` envelope.  The W5 probes synthesize their inputs, so that
manifest has no fixture path.

Build commands are executed directly, never through a shell.  The runner
injects `--offline` and `--locked` only when missing and only before Cargo's
`--` separator, sets `CARGO_TARGET_DIR` to the manifest's explicit target,
rejects Cargo target/config/manifest escapes, forces one Cargo build job and
disabled incremental compilation, and records the effective argv and
environment policy.  Each build target must be a distinct fresh external
directory (absent or empty); no stale target is deleted.  The runner uses
fixed decimal bounds: a 19,200,000,000-byte preflight ceiling, a
20,000,000,000-byte hard target ceiling, and a 20,000,000,000-byte child
address-space cap.  These values cannot be weakened in the manifest.

A build receipt records the observed Git HEAD/tree/status, lock hashes,
probe/fixture hashes, resolved source/target/output paths, target sizes, and
every binary SHA-256.  Before ABBA, the current manifest/runner hashes,
receipt phase/lane/limits, complete event-journal hash/count/order, and all
resolved identities are revalidated.  ABBA refuses to start if any identity or
binary hash differs from the build receipt.

Each child receives a separate stdout/stderr log and a journal event.  Nonzero
exit, empty required output, missing declared output, source mutation, target
growth beyond 20,000,000,000 bytes, post-exit descendants, an unstable target,
or an unrelated Cargo/rustc/rustdoc/rustfmt/clippy or profiler process under
any `/home/zhuhe/code/litchi*` worktree aborts the run.
The process guard follows the launched child subtree, so its own compiler
descendants are allowed while another Litchi worktree's build is rejected.
It identifies the actual executable/comm/cwd; a Python command containing the
word `cargo` is not treated as a compiler.  Failed events and logs remain on
disk and the failed receipt is written before the command returns nonzero.

The build phase requires a completely fresh empty output directory.  The
ABBA phase accepts only the exact build handoff (`commands/`, the build event
journal, and the build receipt), so stale `source-*`/`ops-*` files cannot be
picked up by a historical analyzer glob.

The receipts are orchestration evidence only.  They do not assert correctness
or performance.  The manifests route probe/sheet/source/ops stdout to the
historical `probe-*`, `sheet-*`, `source-*`, and `ops-*` filenames, retain CRUD
JSON under its historical name, and run the real lane `analyze.py` as a
fail-closed validator producing `summary.json`; W5's adapter preserves the
analyzer's full parsed result in its `analysis` field.  A nonempty stdout
stream alone is never accepted.

## W5 OLE owner/history qualification

The standalone OLE owner protocol is `manifests/ole-owner.json`,
`prepare_ole_owner_inputs.py`, and `ole_owner_paired_runner.py`.  Preparation
is source-only: it requires a clean `0ce4bb8b61f03a1e5497633f51b954e76f98aa8c`
baseline, a clean coordinator-selected candidate commit, the immutable
`ad4c52c01333fbc13c7b3cca18e2c0a0d2ccf89d5c07972e2e453fa9bd8d713a` probe, the
reviewed explicit-CFB-limits candidate adapter, the fixed clone probe, and the
fixed W5 lock.  It writes two external Cargo drivers; no production file is
copied into the baseline tree.  The candidate commit is deliberately supplied
at preparation time so this protocol cannot silently qualify a historical
candidate.

The preparation command (all paths are external to both source worktrees) is:

```sh
python3 tools/native_qualification/prepare_ole_owner_inputs.py \
  --baseline-root "$OLE_BEFORE_SOURCE" \
  --candidate-root "$OLE_AFTER_SOURCE" \
  --candidate-ref "$OLE_AFTER_REF" \
  --reviewed-probe /external/inputs/ole_w5_probe-ad4c.rs \
  --clone-probe /external/inputs/ole_package_clone_probe.rs \
  --oracle-probe tools/perf-baseline/examples/ole_owner_oracle.rs \
  --lock /external/inputs/Cargo.lock \
  --drivers-root /external/native-final/ole-owner-drivers \
  --output /external/native-final/ole-owner-preparation.json
```

After coordinator review of that receipt, the build phase creates four fresh
normal/allocator targets (each also builds the correctness oracle) and records
every Cargo event, source identity, lock, driver input, toolchain and executable
hash.  Preparation checks the selected source's actual one-argument baseline
CFB API versus the candidate's two-argument explicit-limits API; the historical
artifact labels and probe hashes are not accepted on their own.  The ABBA phase calibrates all
eleven routes and both clone sizes to a 500 ms target, rejects every leg below
250 ms, runs five `before,after,after,before` blocks on a pinned CPU, and then
runs the paired non-claimable allocator capture for ordinary property reads and
edits, XLS cells/staging, and the 128-step DOC/XLS retained histories.  It
checks host quietness before/after and every 100 ms during each child and keeps
the fixed 19.2 GB preflight, 20 GB target and 20 GB address-space limits:

```sh
python3 tools/native_qualification/ole_owner_paired_runner.py \
  --manifest tools/native_qualification/manifests/ole-owner.json \
  --preparation /external/native-final/ole-owner-preparation.json \
  --output /external/native-final/ole-owner-run1 \
  --target-root /external/native-final/ole-owner-targets/run1 \
  --phase build --cpu "$CPU"

python3 tools/native_qualification/ole_owner_paired_runner.py \
  --manifest tools/native_qualification/manifests/ole-owner.json \
  --preparation /external/native-final/ole-owner-preparation.json \
  --output /external/native-final/ole-owner-run1 \
  --target-root /external/native-final/ole-owner-targets/run1 \
  --phase abba --cpu "$CPU"
```

The measurements and receipts remain external evidence.  The runner's output
is evidence for coordinator review, not a performance acceptance claim.  It
runs `ole_owner_oracle` before calibration, outside every timed loop.  The
oracle's complete source/no-op/replacement OLE bytes, all 64 untouched sibling
metadata and payloads, exact property values, explicit DOC revision metadata,
and explicit XLS cell values are hashed independently and compared across both
endpoints.  Its instrumented consumer variant records alloc/dealloc/realloc
old/new bytes, live/peak state, and the actual retained 129-package/128-patch
history; the normal oracle remains uninstrumented.

The independent post-run validator consumes the preparation, build and ABBA
receipts, journals, logs, immutable build-receipt hash, oracle files and
measurements without invoking Cargo or runner parsing helpers:

```sh
python3 tools/native_qualification/validate_ole_owner_run.py \
  --manifest tools/native_qualification/manifests/ole-owner.json \
  --preparation /external/native-final/ole-owner-preparation.json \
  --build-receipt /external/native-final/ole-owner-run1/build-receipt.json \
  --abba-receipt /external/native-final/ole-owner-run1/abba-receipt.json \
  --measurements /external/native-final/ole-owner-run1/measurements.json \
  --runner tools/native_qualification/ole_owner_paired_runner.py
```

## DrawingML raw protocol

`drawingml_raw_paired_runner.py` and `manifests/drawingml-raw.json` add five
source-fenced raw profiles plus a separate theme-family route and four
detached-owner profiles: chart, diagram, text/media, theme/InkML bulk, and
geometry/color/formatcode2/transform. Every endpoint is materialized into this
generic runner, so it retains the same source/lock/probe identity, Cargo
path, target/address-space, process-quietness, fresh-output, and ABBA guards.
Run `prepare_drawingml_raw_inputs.py` once for each historical parent before
the build phase; it copies only the reviewed adapter paths and preserves the
historical 52ed workspace lock on each recorded baseline and the separately
supplied final-candidate workspace-lock SHA (4100 after source freeze); neither
lock is copied over the other. The perf-baseline lock is independently hashed
as well.

The checked-in probe is an explicit adapter boundary while the raw and owner
API branches converge. It calls the production raw codecs and detached
Package/Edit owners directly; owner before legs use the common pre-owner
945eb1d6 source with the legacy adapter branch. `drawingml_raw_validator.py` rejects a
short leg, altered sample shape, missing source or fixture identity, an
allocator elapsed claim, or a qualified row without outside-clock fidelity
and typed-refusal receipts. Theme bulk and theme-family use separate exact
parents (996651b6 and 11728db); the family route additionally records actual
complete-part `add_family`/`replace_family`/`remove_family` calls under its
typed limit policy. Owner routes never reuse a raw baseline.
It never promotes an absolute current-source profile into a
candidate-versus-baseline claim.

The current-source policy is versioned separately in
`drawingml_raw_paired_runner_v2.py` and `manifests/drawingml-raw-v2.json`.
It imports `prepare_drawingml_raw_inputs_v2.py`, checks all three role locks,
and passes the v2 preparation receipt through
`drawingml_raw_validator_v2.py`; the v1 runner and receipts are never
relabelled.  Its candidate binding is deliberately pending until the
coordinator freezes the exact source commit, manifests, and vendor-tree
digest.  The v2 compile adapter applies the same fail-closed identity rule to
source-only compile receipts.
