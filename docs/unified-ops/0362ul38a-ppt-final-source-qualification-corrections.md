# 0362ul38a — Corrected PPT final-source qualification packet

Date: 2026-10-24. This is a source-only correction to the execution
instructions in `076c8946e1d4`. The earlier record remains retained as failed
handoff provenance; its commands are not execution-ready and must not be used
as an acceptance receipt. This packet does not build Cargo, invoke the probe,
run heaptrack, or run native timing. It keeps the 0357b/0362g/0362j workload,
the tracked 0362g runner and the frozen 28 lazy routes unchanged.

## Fixed source and guard inputs

The candidate source is the clean final-source checkout, not the historical
c0 checkout:

~~~text
source: /home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
HEAD:   472b6705475b073e7fe08f6e11ee17fb1e29f838
tree:   ebc1cee05f124c0c63977aff07a23288c19e7d00
probe:  crates/litchi-ppt/examples/scoped_patch_probe.rs
probe SHA-256: 0bd4d9c1bcad6319968024ff603ae77195810174e496d5394fbae1838021e549
correctness companion: crates/litchi-ppt/examples/scoped_patch_correctness.rs
correctness companion SHA-256: e787342ed25842dae3743d474d3589c2d24bbad2ccb561b3dda5860791c07f60
fixture: test-data/unified-ops/0294/video.ppt
fixture bytes/SHA-256: 43520 / f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a
~~~

The accepted build setup is the pinned gate-local Cargo wrapper, not the old
`.coord/w8-file/env.sh` environment file. The reviewed target/process
supervisor is the tracked `0362ul29` `run_safe_gate.py`; its selected PPT step
must bind these wrapper paths and hashes. The shell must not replace them
through an inherited `PATH`:

~~~text
wrapper: docs/unified-ops/results/0362ul29/tooling/bin/cargo
wrapper SHA-256: 495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d
inner memory-capped shim: /home/zhuhe/code/litchi-worktrees/.coord/bin/cargo
inner shim SHA-256: 1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8
target/process supervisor: docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py
supervisor SHA-256: b8603b469af213b097fcf1111ce57dbe2c8e6e3de19a9bbe01514f53eb6e663a
reference compile command: /home/zhuhe/code/litchi-worktrees/.coord/rtf-fuzz-39a417-20261024/compile-command.txt
reference compile command SHA-256: 81ab917ba0447718932a08dcb5e0ef1062fc0bda41d2aa4df48a5127019807b1
reference compile receipt: /home/zhuhe/code/litchi-worktrees/.coord/rtf-fuzz-39a417-20261024/receipt.json
reference compile receipt SHA-256: ef1fc17ff1e95bab311bacc130e06dfeb292fcf86090a3d61c5071bc6b4c4f3a
~~~

The existing reviewed outer process boundary is also retained rather than
reimplemented for PPT:

~~~text
boundary: /home/zhuhe/code/litchi-worktrees/uo2-docx-ink-final-handoff/tools/ink_safe_runner.py
bytes/SHA-256: 17630 / 1bc982f4495a16c30ac4ca2779351aec1e420781a6d6cb7c2b1f33ccff88fede
host-process adapter: docs/unified-ops/results/0362g/ppt_host_process_guard.py
adapter SHA-256: 08c73174304061beeaf778df763be3507befd8200db72a8617a5d8ddc42e12b8
correctness receipt validator: docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py
validator SHA-256: db0ac0f330db5c15334f318efb8e4152d7752ac0731d7cf83f8b4d5e1ddfa95e
candidate lock guard: docs/unified-ops/results/0362g/validate_ppt_candidate_lock.py
candidate lock guard SHA-256: 5ed31f6bf3377b1936b58e214fe2ea04c796b9b9f9508079a7b02a06820ab714
final-source supervisor/wrapper hashes: must be freshly rebound after
0362ul39/source freeze
~~~

The accepted `run_safe_gate.py` binds the target and process policy: target
preflight at 18,000,000,000 bytes, live guard at 19,000,000,000 bytes, the
20,000,000,000-byte hard limit, post-exit stabilization, one Cargo job, 20-GB
memory, zero swap and owned-process guards. Its tracked wrapper supplies the
offline Cargo environment. The outer Ink boundary remains generic argv
infrastructure and binds a systemd
scope to `MemoryMax=20,000,000,000`, zero swap, a phase deadline, `TasksMax`,
and `KillMode=control-group`; its child sets `RLIMIT_FSIZE=1,000,000,000`,
disables core dumps, polls aggregate output and TMP trees, and kills the
owned process group on a deadline or cap breach. The host-process adapter runs
the unchanged 0362g runner as its owned child and recognizes both
`heaptrack`, `heaptrack_inter` and `heaptrack_print` (as well as the existing
compiler/profiler names) before, during and after the child. The historical
0362g runner remains the measured inner process guard and retains its
phase-span checks and inverse byte oracle; the adapter does not change its
argv or timing loop.

## Dependency-lock policy

The protected historical original and matched-Sip endpoints were built with
lock SHA `40dcbcff741065997d9bc583140a22d921ebe2b4f321570af2587d1cf22d3b5c`.
The final candidate source contains lock SHA
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.
These are intentionally not relabelled as one lock. The candidate keeps its
4100 lock; the protected references are not rebuilt or overlaid.

The relevant parity is concrete: the two locks have equal
`(name, version, source, checksum)` package identity tuples, and the
`litchi-ppt` package dependency stanza is identical. The only dependency-edge
stanza differences are the four non-PPT workspace packages introduced by the
assembled final source: `litchi-iwa-core`, `litchi-ops`, `litchi-xlsb` and
`litchi-xlsx`. A candidate receipt must record both lock hashes and this
comparison. Any registry identity or `litchi-ppt` closure difference is a
hard stop; no network resolution, lock rewrite or old-lock copy is allowed.

Run this bounded, read-only policy check before the candidate build. Its
receipt intentionally lives in a separate fresh directory, so the later
candidate `BUILD`, `TARGET` and build-TMP paths remain absent and can still
pass their own fresh-build precondition. It writes only the fresh receipt:

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
LOCK_RECEIPT="$ROOT/lock-policy"
REFERENCE_LOCK=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-admission-20261021/Cargo.lock
CANDIDATE_LOCK="$SOURCE/Cargo.lock"
test ! -e "$LOCK_RECEIPT"
mkdir -p "$LOCK_RECEIPT"
python3 - "$REFERENCE_LOCK" "$CANDIDATE_LOCK" "$LOCK_RECEIPT/lock-policy.json" <<'PY'
import hashlib
import json
from pathlib import Path
import sys
import tomllib

reference, candidate, receipt = map(Path, sys.argv[1:])
expected_reference = "40dcbcff741065997d9bc583140a22d921ebe2b4f321570af2587d1cf22d3b5c"
expected_candidate = "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1"

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

if digest(reference) != expected_reference or digest(candidate) != expected_candidate:
    raise SystemExit("lock hash is not the reviewed reference or final candidate")

def identity(package):
    return tuple(package.get(key) for key in ("name", "version", "source", "checksum"))

def load(path):
    return tomllib.loads(path.read_text())

ref_packages = load(reference)["package"]
candidate_packages = load(candidate)["package"]
ref_ids = sorted(identity(package) for package in ref_packages)
candidate_ids = sorted(identity(package) for package in candidate_packages)
if ref_ids != candidate_ids:
    raise SystemExit("registry/package identity tuples differ")

def by_identity(packages):
    return {identity(package): package for package in packages}

ref_by_id = by_identity(ref_packages)
candidate_by_id = by_identity(candidate_packages)
changed = sorted(
    ref_by_id[item]["name"]
    for item in ref_by_id
    if ref_by_id[item].get("dependencies", []) != candidate_by_id[item].get("dependencies", [])
)
expected_changed = ["litchi-iwa-core", "litchi-ops", "litchi-xlsb", "litchi-xlsx"]
if changed != expected_changed:
    raise SystemExit(f"unexpected dependency-edge delta: {changed!r}")

def named(packages, name):
    values = [package for package in packages if package["name"] == name]
    if len(values) != 1:
        raise SystemExit(f"expected one {name} package, found {len(values)}")
    return values[0]

ref_ppt = named(ref_packages, "litchi-ppt")
candidate_ppt = named(candidate_packages, "litchi-ppt")
if ref_ppt.get("dependencies", []) != candidate_ppt.get("dependencies", []):
    raise SystemExit("litchi-ppt dependency closure differs")

receipt.write_text(json.dumps({
    "reference_lock": {"path": str(reference), "sha256": digest(reference)},
    "candidate_lock": {"path": str(candidate), "sha256": digest(candidate)},
    "package_identity_tuples_equal": True,
    "litchi_ppt_dependencies_equal": True,
    "known_dependency_edge_delta_packages": changed,
}, indent=2) + "\n")
PY
~~~

## Corrected guarded build and probe oracle

The earlier direct Cargo blocks are retired. The candidate build and the seven
correctness commands must run through the accepted `run_safe_gate.py` step,
which records source/lock identity and enforces target preflight, live target
checks, the hard target limit, post-exit stabilization, one Cargo job, 20-GB
memory and process guards. No `env.sh` or direct Cargo invocation may replace
that supervisor.

This leaf registers four source-bound PPT steps in the existing
`run_safe_gate.py` plan: `ppt-scoped-build`,
`ppt-scoped-correctness-suite`, `ppt-scoped-correctness`, and
`ppt-scoped-correctness-validate`. The final assembled source must retain and
rebind them. They are auxiliary steps: an ordinary run without an explicit
`--only` or `--start-at` selection retains the existing full-gate plan and
count (53 steps with the complete operations feature manifest; 57 only when
all four auxiliaries are explicitly composed). The future command shape (with
final-source hashes filled by the rebind receipt) is:

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
SUPERVISOR="$SOURCE/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py"
RUN_ROOT=/external/ppt-final-source-472b670547
STATE="$RUN_ROOT/ppt-state"
TARGET="$RUN_ROOT/target"
PPT_STEPS=ppt-scoped-build,ppt-scoped-correctness-suite,ppt-scoped-correctness,ppt-scoped-correctness-validate
test ! -e "$RUN_ROOT"
mkdir -p "$RUN_ROOT"
test -f "$SUPERVISOR"; test -d "$(dirname "$STATE")"; test ! -e "$TARGET"
test "$(sha256sum "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" | cut -d' ' -f1)" = 0bd4d9c1bcad6319968024ff603ae77195810174e496d5394fbae1838021e549
unset LITCHI_CAPPED LITCHI_BLESS
python3 "$SUPERVISOR" --root "$SOURCE" --target "$TARGET" --state "$STATE" --only "$PPT_STEPS"
~~~

The coordinator-selected step/state must bind the final source HEAD/tree,
`4100…` lock, probe and fixture, fresh external target/output/TMP paths, and
the exact release probe artifact. Its correctness command list must retain
these unchanged argv selectors with `RUST_TEST_THREADS=1` and
`--test-threads=1`:

Both Cargo-bearing auxiliary steps first run the candidate lock guard against
the source `Cargo.lock` and the approved candidate SHA
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`. A wrong
lock exits before the host guard or Cargo is launched; the guard is therefore
part of the preflight, not a post-build observation.

~~~text
export RUST_TEST_THREADS=1
cargo test --offline --locked -p litchi-ppt --all-features --lib scoped::tests -- --test-threads=1
cargo test --offline --locked -p litchi-ppt --all-features --lib -- --test-threads=1
cargo test --offline --locked -p litchi-cfb --all-features --lib -- --test-threads=1
cargo test --offline --locked -p litchi-ppt --all-features --tests -- --test-threads=1
cargo test --offline --locked -p litchi-cfb --all-features --tests -- --test-threads=1
cargo clippy --offline --locked -p litchi-ppt --all-features --all-targets -- -D warnings
cargo clippy --offline --locked -p litchi-cfb --all-features --all-targets -- -D warnings
~~~

The historical probe does not emit an `inverse_verified` JSON field and remains
unchanged for timing comparability. The separate
`scoped_patch_correctness` companion uses the same source fixture and public
`edit_external_media().set_path(1, Some("clip-0000.avi"))` operation outside
the timed workload. It parses the committed target through
`slide_order::Snapshot`, rejects unchanged target bytes or an unexpected path,
replays the patch and inverse, compares the forward replay bytes to the
committed target bytes, and writes source/target/replay/inverse hashes to an
outside-clock receipt. The validator binds that receipt to the selected
fixture, rejects equal source/target hashes and any path other than
`clip-0000.avi`, requires the forward replay hash and boolean to match the
committed target, and requires the inverse hash to equal the source hash.
The companion and validator are separate gate steps, so no checksum or extra
serialization is added to the historical timed loop:

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
RUN_ROOT=/external/ppt-final-source-472b670547
TARGET="$RUN_ROOT/target"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
CORRECTNESS="$TARGET/release/examples/scoped_patch_correctness"
RECEIPT="$RUN_ROOT/ppt-state/ppt-scoped-correctness.json"
VALIDATOR="$SOURCE/docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"
test -x "$CORRECTNESS"; test -f "$VALIDATOR"; test -f "$FIXTURE"
test -f "$RECEIPT"
"$CORRECTNESS" "$FIXTURE" "$RECEIPT"
python3 "$VALIDATOR" "$RECEIPT" "$FIXTURE"
~~~

The historical 0362j calibration counts are distinct and remain historical
inputs, not fixed commands for a fresh native run:

| endpoint | `N` | `FIRST` | `REPLAY` |
| --- | ---: | ---: | ---: |
| original reference | 7400 | 3500 | 42800 |
| matched-Sip reference | 7400 | 3300 | 42600 |

The fresh 0362g runner calibrates both endpoints itself. Its native commands
must not hard-code either historical row; `--blocks 5` and the accepted runner
are the only workload controls in the fresh timing packet.

## Corrected bounded native commands

The following uses the existing outer boundary around the unchanged 0362g
runner. The one-gigabyte output/TMP caps comfortably exceed the retained
0362j full-workload artifact sizes (largest recorded compressed heap artifact
2,603,157 bytes and massif artifact 259,644 bytes); if a fresh output or TMP
tree grows past the cap, the boundary fails closed rather than accepting a
partial result. The native output paths, TMP paths and outer logs are all
fresh and pairwise separate.

~~~sh
set -eu
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
BOUNDARY=/home/zhuhe/code/litchi-worktrees/uo2-docx-ink-final-handoff/tools/ink_safe_runner.py
HOST_GUARD="$SOURCE/docs/unified-ops/results/0362g/ppt_host_process_guard.py"
RUNNER="$SOURCE/docs/unified-ops/results/0362g/ppt_phase_abba.py"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
AFTER="$ROOT/candidate-build/protected/scoped_patch_probe-472b670547"
ORIGINAL=/home/zhuhe/code/litchi-worktrees/.coord/ppt-instrumented-baselines-20261021/original-e673d4addb-scoped_patch_probe
SIP=/home/zhuhe/code/litchi-worktrees/.coord/ppt-sip-common-lock-20261021/sip-common-lock-scoped_patch_probe
test "$(sha256sum "$BOUNDARY" | cut -d' ' -f1)" = 1bc982f4495a16c30ac4ca2779351aec1e420781a6d6cb7c2b1f33ccff88fede
test "$(sha256sum "$HOST_GUARD" | cut -d' ' -f1)" = 08c73174304061beeaf778df763be3507befd8200db72a8617a5d8ddc42e12b8
test "$(sha256sum "$RUNNER" | cut -d' ' -f1)" = daf58fc8f54ecc30b105f934499ecb313975ebcbfe7a6410356afe6ab699a75c
test "$(sha256sum "$ORIGINAL" | cut -d' ' -f1)" = 9f23a1a9119cfb541361e7661c3384a1b2c63c59ae9d95f4842f4a9c1d5a5b28
test "$(sha256sum "$SIP" | cut -d' ' -f1)" = 61272a3cc329c27b8ad0dfa1c20934e4c95c6e7b0d57a0de8331bae868b01139
test -x "$AFTER"; test -x "$ORIGINAL"; test -x "$SIP"
mkdir -p "$ROOT"
for path in "$ROOT/native-original" "$ROOT/native-sip" "$ROOT/tmp-native-original" "$ROOT/tmp-native-sip" "$ROOT/outer-native-original" "$ROOT/outer-native-sip"; do test ! -e "$path"; done

python3 "$BOUNDARY" abba \
  --output "$ROOT/native-original" \
  --tmp "$ROOT/tmp-native-original" \
  --log-dir "$ROOT/outer-native-original" \
  --unit litchi-ppt-final-472b-original-a1 \
  --deadline 7200 --output-cap 1000000000 --tmp-cap 1000000000 -- \
  python3 "$HOST_GUARD" -- python3 "$RUNNER" \
  --before "$ORIGINAL" --after "$AFTER" --fixture "$FIXTURE" \
  --out "$ROOT/native-original" \
  --before-commit e673d4addbf4b4224f377e419702a33b3d6958fe \
  --after-commit 472b6705475b073e7fe08f6e11ee17fb1e29f838 \
  --harness "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" \
  --cpu 29 --blocks 5

python3 "$BOUNDARY" abba \
  --output "$ROOT/native-sip" \
  --tmp "$ROOT/tmp-native-sip" \
  --log-dir "$ROOT/outer-native-sip" \
  --unit litchi-ppt-final-472b-sip-a1 \
  --deadline 7200 --output-cap 1000000000 --tmp-cap 1000000000 -- \
  python3 "$HOST_GUARD" -- python3 "$RUNNER" \
  --before "$SIP" --after "$AFTER" --fixture "$FIXTURE" \
  --out "$ROOT/native-sip" \
  --before-commit e4bdab60667470124c4da202477ba6bb1734f256 \
  --after-commit 472b6705475b073e7fe08f6e11ee17fb1e29f838 \
  --harness "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" \
  --cpu 29 --blocks 5
~~~

No timing or performance claim is valid until both boundary receipts and all
0362g per-leg host/output/time artifacts are complete, the coordinator grants
quiet CPU 29, and an independent review confirms no boundary violation.

## Corrected bounded allocation command

The old `full_heap.py` is not reused: it hard-codes the c0 source, old output
tree and historical working directory. The candidate allocation leg uses the
matched-Sip full-workload count for the candidate/reference allocation
comparison (`7400 3300 42600`) and the same outer boundary. The historical
0362j full-workload rows are allocation/heap evidence only; they are not RSS
or latency evidence.

~~~sh
set -eu
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
BOUNDARY=/home/zhuhe/code/litchi-worktrees/uo2-docx-ink-final-handoff/tools/ink_safe_runner.py
HOST_GUARD="$SOURCE/docs/unified-ops/results/0362g/ppt_host_process_guard.py"
BIN="$ROOT/candidate-build/protected/scoped_patch_probe-472b670547"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
ALLOC="$ROOT/candidate-allocation"
TMP="$ROOT/tmp-candidate-allocation"
LOG="$ROOT/outer-candidate-allocation"
ANALYSIS="$ROOT/candidate-heap-analysis"
ANALYSIS_TMP="$ROOT/tmp-candidate-heap-analysis"
ANALYSIS_LOG="$ROOT/outer-candidate-heap-analysis"
test "$(sha256sum "$BOUNDARY" | cut -d' ' -f1)" = 1bc982f4495a16c30ac4ca2779351aec1e420781a6d6cb7c2b1f33ccff88fede
test "$(sha256sum "$HOST_GUARD" | cut -d' ' -f1)" = 08c73174304061beeaf778df763be3507befd8200db72a8617a5d8ddc42e12b8
test -x "$BIN"; test -f "$FIXTURE"
for path in "$ALLOC" "$TMP" "$LOG" "$ANALYSIS" "$ANALYSIS_TMP" "$ANALYSIS_LOG"; do test ! -e "$path"; done
mkdir -p "$ALLOC" "$ANALYSIS"

python3 "$BOUNDARY" allocation \
  --output "$ALLOC" --tmp "$TMP" --log-dir "$LOG" \
  --unit litchi-ppt-final-472b-heap-a1 \
  --deadline 1800 --output-cap 1000000000 --tmp-cap 1000000000 -- \
  python3 "$HOST_GUARD" -- \
  heaptrack --record-only -o "$ALLOC/candidate" "$BIN" "$FIXTURE" 7400 3300 42600

python3 "$BOUNDARY" allocation \
  --output "$ANALYSIS" --tmp "$ANALYSIS_TMP" --log-dir "$ANALYSIS_LOG" \
  --unit litchi-ppt-final-472b-heap-analysis-a1 \
  --deadline 1800 --output-cap 1000000000 --tmp-cap 1000000000 -- \
  python3 "$HOST_GUARD" -- /bin/sh -c "heaptrack_print -f '$ALLOC/candidate.zst' --print-peaks=0 --print-allocators=0 --print-temporary=0 --print-massif '$ANALYSIS/candidate.massif' > '$ANALYSIS/candidate.analysis' 2> '$ANALYSIS/candidate.analysis.stderr'"
~~~

The allocation receipt must bind the candidate binary and fixture hashes,
counts, outer receipts, heaptrack output, analysis output, exit codes and
boundary caps. It remains allocation/heap evidence only. The two historical
protected references and their 0362j allocation facts remain unchanged.

## Remaining blockers

1. The final assembled source must retain the companion, validator, host
   adapter and four `run_safe_gate.py` steps, then rebind the source,
   supervisor, step/state, probe, fixture and tool hashes. The current 472b
   checkout does not yet satisfy final-source bindings.
2. The candidate build, current-source correctness/preflight, bounded native
   ABBA and bounded allocation receipts do not yet exist. The registered PPT
   steps must exercise the companion and seven correctness
   selectors with
   `RUST_TEST_THREADS=1`, `--test-threads=1`, target pre/live/post checks and
   fresh output/TMP paths.
3. The lock policy check must pass with candidate `4100…` and references
   `40dcb…`; neither lock may be silently substituted. CPU 29 still needs a
   quiet coordinator grant, and both native boundary receipts need independent
   review.
4. This is preparation only. It makes no native, allocation, fidelity,
   performance or final-gate acceptance claim. Memo item 22 remains the only
   deferred audit.

## Coordinator composition (2026-10-25)

The forward-replay correction in `6242efca31952b1755f86befb5ee82e27b10e0f4`
requires the replayed bytes to equal the committed target before checking
inverse restoration. The receipt validator rejects mismatched replay hashes
and false replay status. This prevents a pair of no-op apply operations from
satisfying the inverse-only oracle. These checks remain outside the timed
workload.

The composed tooling row explicitly registers this packet's three Python
tests and `tools.test_w8_quiet_abba`. The projection-policy suite no longer
imports the latter's test class for implicit discovery, avoiding duplicate
execution. All 44 projection, lifecycle and PPT wiring tests pass together.
The normal gate remains 53 steps; its four PPT qualification steps are
opt-in. This source/fake verification does not replace the pending full
tooling row or actual qualification runs.
