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
fixture: test-data/unified-ops/0294/video.ppt
fixture bytes/SHA-256: 43520 / f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a
~~~

The build must use the pinned gate-local Cargo wrapper, not the old
`.coord/w8-file/env.sh` environment file:

~~~text
wrapper: /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin/cargo
wrapper SHA-256: 495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d
inner memory-capped shim: /home/zhuhe/code/litchi-worktrees/.coord/bin/cargo
inner shim SHA-256: 1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8
~~~

The existing reviewed outer process boundary is also retained rather than
reimplemented for PPT:

~~~text
boundary: /home/zhuhe/code/litchi-worktrees/uo2-docx-ink-final-handoff/tools/ink_safe_runner.py
bytes/SHA-256: 16835 / f184664f29c1504e9d070931dcb476be2ac680fb5d7c5d62785f6fcdb41b4c16
~~~

Its accepted boundary is generic argv infrastructure. It binds a systemd
scope to `MemoryMax=20,000,000,000`, zero swap, a phase deadline, `TasksMax`,
and `KillMode=control-group`; its child sets `RLIMIT_FSIZE=1,000,000,000`,
disables core dumps, polls aggregate output and TMP trees, and kills the
owned process group on a deadline or cap breach. The 0362g runner remains the
inner native process guard and retains its before/during/after quiet checks,
descendant exclusion, phase-span checks and inverse byte oracle.

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

Run this bounded, read-only policy check after creating the fresh build
directory and before the candidate build. It writes only the fresh receipt:

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
BUILD=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c/candidate-build
REFERENCE_LOCK=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-admission-20261021/Cargo.lock
CANDIDATE_LOCK="$SOURCE/Cargo.lock"
python3 - "$REFERENCE_LOCK" "$CANDIDATE_LOCK" "$BUILD/lock-policy.json" <<'PY'
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

The build shell must unset inherited qualification switches, bind one Cargo
job and the decimal 20-GB cap, create a fresh TMP tree, and change directory
to the final source before the first Cargo invocation. No `env.sh` is sourced.

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
TARGET="$ROOT/candidate-target"
BUILD="$ROOT/candidate-build"
TMP="$ROOT/candidate-build-tmp"
PROTECTED="$BUILD/protected/scoped_patch_probe-472b670547"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
CARGO_SHIM=/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin/cargo
INNER_SHIM=/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo
test "$(sha256sum "$CARGO_SHIM" | cut -d' ' -f1)" = 495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d
test "$(sha256sum "$INNER_SHIM" | cut -d' ' -f1)" = 1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8
test "$(git -C "$SOURCE" rev-parse HEAD)" = 472b6705475b073e7fe08f6e11ee17fb1e29f838
test "$(git -C "$SOURCE" rev-parse HEAD^{tree})" = ebc1cee05f124c0c63977aff07a23288c19e7d00
test -z "$(git -C "$SOURCE" status --short)"
test "$(sha256sum "$SOURCE/Cargo.lock" | cut -d' ' -f1)" = 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
test "$(sha256sum "$SOURCE/rust-toolchain.toml" | cut -d' ' -f1)" = e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
test "$(sha256sum "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" | cut -d' ' -f1)" = 0bd4d9c1bcad6319968024ff603ae77195810174e496d5394fbae1838021e549
test "$(sha256sum "$SOURCE/test-data/unified-ops/0294/video.ppt" | cut -d' ' -f1)" = f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a
for path in "$TARGET" "$BUILD" "$TMP"; do test ! -e "$path"; done
mkdir -p "$BUILD" "$TMP"
unset LITCHI_CAPPED LITCHI_BLESS
export PATH="$(dirname "$CARGO_SHIM"):/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH"
export CARGO_TARGET_DIR="$TARGET" LITCHI_GATE_TARGET_ROOT="$TARGET"
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G
export CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_TERM_COLOR=never
export TMPDIR="$TMP"
cd "$SOURCE"
cargo build --offline --locked --release -p litchi-ppt --example scoped_patch_probe \
  >"$BUILD/build.log" 2>&1
test -x "$TARGET/release/examples/scoped_patch_probe"
install -D -m 0755 "$TARGET/release/examples/scoped_patch_probe" "$PROTECTED"
sha256sum "$PROTECTED" >"$BUILD/candidate-binary.sha256"
~~~

Run the existing 0362j correctness selection with the same explicit
environment, using a separate fresh TMP tree but the candidate target. The
seven Cargo commands are unchanged from the accepted selection.

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
TARGET="$ROOT/candidate-target"
TMP="$ROOT/candidate-check-tmp"
CARGO_SHIM=/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-drawingml-20261024/bin/cargo
test ! -e "$TMP"
mkdir -p "$TMP"
unset LITCHI_CAPPED LITCHI_BLESS
export PATH="$(dirname "$CARGO_SHIM"):/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH"
export CARGO_TARGET_DIR="$TARGET" LITCHI_GATE_TARGET_ROOT="$TARGET"
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G
export CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_TERM_COLOR=never RUST_TEST_THREADS=2
export TMPDIR="$TMP"
cd "$SOURCE"
cargo test --offline --locked -p litchi-ppt --all-features --lib scoped::tests -- --test-threads=2
cargo test --offline --locked -p litchi-ppt --all-features --lib -- --test-threads=2
cargo test --offline --locked -p litchi-cfb --all-features --lib -- --test-threads=2
cargo test --offline --locked -p litchi-ppt --all-features --tests -- --test-threads=2
cargo test --offline --locked -p litchi-cfb --all-features --tests -- --test-threads=2
cargo clippy --offline --locked -p litchi-ppt --all-features --all-targets -- -D warnings
cargo clippy --offline --locked -p litchi-cfb --all-features --all-targets -- -D warnings
~~~

The probe does not emit an `inverse_verified` JSON field. Its actual oracle is
the source check in `scoped_patch_probe.rs`: each replayed patch is inverted,
the restored package is serialized, and the process returns nonzero if those
bytes differ from the source. The preflight therefore checks the real output
keys and the zero exit/stderr, without inventing a field:

~~~sh
set -eu
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
BUILD="$ROOT/candidate-build"
PROTECTED="$BUILD/protected/scoped_patch_probe-472b670547"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
test -x "$PROTECTED"
test ! -e "$BUILD/preflight-1-1-1.json"; test ! -e "$BUILD/preflight-1-1-1.stderr"
"$PROTECTED" "$FIXTURE" 1 1 1 \
  >"$BUILD/preflight-1-1-1.json" 2>"$BUILD/preflight-1-1-1.stderr"
test ! -s "$BUILD/preflight-1-1-1.stderr"
python3 - "$BUILD/preflight-1-1-1.json" "$FIXTURE" <<'PY'
import json
from pathlib import Path
import sys

row = json.loads(Path(sys.argv[1]).read_text())
fixture_size = Path(sys.argv[2]).stat().st_size
for key in ("patches", "first_count", "replay_count", "source_len", "target_len",
            "retained_patch_bytes", "retained_payload_bytes", "retained_block_span_bytes",
            "retained_block_span_count", "first_total_ns", "commit_total_ns",
            "apply_total_ns", "inverse_apply_total_ns"):
    if key not in row:
        raise SystemExit(f"missing probe field: {key}")
if row["patches"] != 1 or row["first_count"] != 1 or row["replay_count"] != 1:
    raise SystemExit("preflight counts differ")
if row["source_len"] != fixture_size or row["target_len"] <= 0:
    raise SystemExit("preflight extents are invalid")
if row["inverse_apply_total_ns"] <= 0:
    raise SystemExit("preflight inverse phase was not measured")
if "inverse_verified" in row:
    raise SystemExit("probe output fabricated the absent inverse_verified field")
PY
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
RUNNER="$SOURCE/docs/unified-ops/results/0362g/ppt_phase_abba.py"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
AFTER="$ROOT/candidate-build/protected/scoped_patch_probe-472b670547"
ORIGINAL=/home/zhuhe/code/litchi-worktrees/.coord/ppt-instrumented-baselines-20261021/original-e673d4addb-scoped_patch_probe
SIP=/home/zhuhe/code/litchi-worktrees/.coord/ppt-sip-common-lock-20261021/sip-common-lock-scoped_patch_probe
test "$(sha256sum "$BOUNDARY" | cut -d' ' -f1)" = f184664f29c1504e9d070931dcb476be2ac680fb5d7c5d62785f6fcdb41b4c16
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
  python3 "$RUNNER" \
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
  python3 "$RUNNER" \
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
BIN="$ROOT/candidate-build/protected/scoped_patch_probe-472b670547"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
ALLOC="$ROOT/candidate-allocation"
TMP="$ROOT/tmp-candidate-allocation"
LOG="$ROOT/outer-candidate-allocation"
ANALYSIS="$ROOT/candidate-heap-analysis"
ANALYSIS_TMP="$ROOT/tmp-candidate-heap-analysis"
ANALYSIS_LOG="$ROOT/outer-candidate-heap-analysis"
test "$(sha256sum "$BOUNDARY" | cut -d' ' -f1)" = f184664f29c1504e9d070931dcb476be2ac680fb5d7c5d62785f6fcdb41b4c16
test -x "$BIN"; test -f "$FIXTURE"
for path in "$ALLOC" "$TMP" "$LOG" "$ANALYSIS" "$ANALYSIS_TMP" "$ANALYSIS_LOG"; do test ! -e "$path"; done
mkdir -p "$ALLOC" "$ANALYSIS"

python3 "$BOUNDARY" allocation \
  --output "$ALLOC" --tmp "$TMP" --log-dir "$LOG" \
  --unit litchi-ppt-final-472b-heap-a1 \
  --deadline 1800 --output-cap 1000000000 --tmp-cap 1000000000 -- \
  heaptrack --record-only -o "$ALLOC/candidate" "$BIN" "$FIXTURE" 7400 3300 42600

python3 "$BOUNDARY" allocation \
  --output "$ANALYSIS" --tmp "$ANALYSIS_TMP" --log-dir "$ANALYSIS_LOG" \
  --unit litchi-ppt-final-472b-heap-analysis-a1 \
  --deadline 1800 --output-cap 1000000000 --tmp-cap 1000000000 -- \
  /bin/sh -c "heaptrack_print -f '$ALLOC/candidate.zst' --print-peaks=0 --print-allocators=0 --print-temporary=0 --print-massif '$ANALYSIS/candidate.massif' > '$ANALYSIS/candidate.analysis' 2> '$ANALYSIS/candidate.analysis.stderr'"
~~~

The allocation receipt must bind the candidate binary and fixture hashes,
counts, outer receipts, heaptrack output, analysis output, exit codes and
boundary caps. It remains allocation/heap evidence only. The two historical
protected references and their 0362j allocation facts remain unchanged.

## Remaining blockers

1. The candidate build, current-source correctness/preflight, bounded native
   ABBA and bounded allocation receipts do not yet exist.
2. The lock policy check must pass with candidate `4100…` and references
   `40dcb…`; neither lock may be silently substituted.
3. CPU 29 still needs a quiet coordinator grant and both native boundary
   receipts need independent review.
4. This is preparation only. It makes no native, allocation, fidelity,
   performance or final-gate acceptance claim. Memo item 22 remains the only
   deferred audit.
