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

## DrawingML raw protocol

`drawingml_raw_paired_runner.py` and `manifests/drawingml-raw.json` add five
source-fenced profiles: chart, diagram, text/media, theme/InkML, and
geometry/color/formatcode2/transform. The profiles are materialized into this
generic runner, so they retain the same source/lock/probe identity, Cargo
path, target/address-space, process-quietness, fresh-output, and ABBA guards.
Run `prepare_drawingml_raw_inputs.py` once for each historical parent before
the build phase; it copies only the reviewed adapter paths and preserves the
52ed historical workspace lock (plus the separately hashed perf-baseline
lock).

The checked-in probe is an explicit adapter boundary while the five API
branches converge. It emits `unavailable` rows and no elapsed values until a
final-source adapter is supplied. `drawingml_raw_validator.py` rejects a
short leg, altered sample shape, missing source or fixture identity, an
allocator elapsed claim, or a qualified row without outside-clock fidelity
and typed-refusal receipts. The combined theme/InkML profile is fail-closed
on the exact theme parent; its separately listed Ink/action parents cannot be
selected as an ambiguous baseline. It never promotes an absolute
current-source profile into a candidate-versus-baseline claim.
