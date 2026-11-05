# 0362ul38 — PPT final-source changed-commit qualification handoff

Date: 2026-10-24. This is a source-only, executable handoff. It does not
build Cargo, invoke the probe, run heaptrack, run native timing, or promote a
historical result. It preserves the 0357b/0362g/0362j workload and uses the
reviewed 0362g runner; it does not introduce a primary_noop workload and it
does not alter the frozen 28 lazy routes.

## Qualification boundary

The final-scope audit identifies a real evidence boundary. The accepted 0362j
native result measured source c0f346607bb28559ae21fa780c835dc70e87fb0b,
while the final production crate tree is the 472b tree
ebc1cee05f124c0c63977aff07a23288c19e7d00. The 3431 sparse leaf has the same
crates subtree (68caf2eff3de127d4774601791277e0d364937be) but a different
full-worktree tree (061c4bc1134fe9599bde4e6f3e0d608538008950). The current
final-source endpoint is the clean external checkout:

~~~
/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
HEAD 472b6705475b073e7fe08f6e11ee17fb1e29f838
tree ebc1cee05f124c0c63977aff07a23288c19e7d00
~~~

The relevant source identities are:

| file | 0362j candidate | final source |
| --- | --- | --- |
| crates/litchi-cfb/src/shared.rs | c98f041fad4d59674536a54fdda1fde052587bd0c1ff10eb40f76ce4d4b6bdd2 | ea777d00ae1fc6dec49e5da0f5e3131f6307cc75401b577d6804db3a8bfe9835 |
| crates/litchi-ppt/src/scoped.rs | 88af7af8c76325590e16a5777388794a814506ec5f9f737d948d5091d238faf3 | 2ea31f4cf1517b50e135d9cb2f2bbf1772a101b99671082f7ce38ab1d1ade461 |
| crates/litchi-ppt/src/slide_order.rs | e81ca2ca441eb95278c5348ee5207191b4c070b0757811f7476b8bd111b3e45a | 65297b805addb9054a9274fc68cb83e7d9ec837d828011d086e5c8cc558e1c9a |
| crates/litchi-ppt/src/scoped/digest.rs | 070a534c47609c12704a85cd0e824323249f93aaba9910920564b7809261fdbd | same |
| crates/litchi-ppt/examples/scoped_patch_probe.rs | 0bd4d9c1bcad6319968024ff603ae77195810174e496d5394fbae1838021e549 | same |

Therefore the old c0 binary, 0362j timing rows, and c0 heap rows are retained
historical evidence only. They cannot be attributed to the final source.

## Workload and retained inputs

The workload remains the accepted scoped changed-commit/replay workload:

* open video.ppt outside the timed region;
* first commit: a fresh Package per iteration and
  edit_external_media().set_path(1, Some("first-####.avi"));
* steady commit plus publication: one source package and
  set_path(1, Some("clip-####.avi")), retaining every patch;
* forward replay and inverse replay, with inverse output checked byte-for-byte
  against the source outside the timed phase;
* five complete before, after, after, before blocks for each reference;
* calibration is performed independently by the runner, with the historical
  0362j workload sizes retained (N=7400, FIRST=3300 for the original
  reference and FIRST=3200 for matched Sip; REPLAY=42600), subject to the
  runner's endpoint calibration and its 20,000/20,000/200,000 ceilings.

The fixture is present and unchanged:

~~~
/home/zhuhe/code/litchi-worktrees/.coord/ppt-replay-c0-source-20261021/test-data/unified-ops/0294/video.ppt
bytes 43520
sha256 f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a
~~~

The tracked probe is the same 7,385-byte source as 0362j:

~~~
crates/litchi-ppt/examples/scoped_patch_probe.rs
sha256 0bd4d9c1bcad6319968024ff603ae77195810174e496d5394fbae1838021e549
~~~

The protected reference binaries are retained, not rebuilt or modified:

| reference | path | bytes | SHA-256 |
| --- | --- | ---: | --- |
| original | /home/zhuhe/code/litchi-worktrees/.coord/ppt-instrumented-baselines-20261021/original-e673d4addb-scoped_patch_probe | 1,716,976 | 9f23a1a9119cfb541361e7661c3384a1b2c63c59ae9d95f4842f4a9c1d5a5b28 |
| matched SipHash | /home/zhuhe/code/litchi-worktrees/.coord/ppt-sip-common-lock-20261021/sip-common-lock-scoped_patch_probe | 1,770,160 | 61272a3cc329c27b8ad0dfa1c20934e4c95c6e7b0d57a0de8331bae868b01139 |

The reviewed runner is the tracked 0362g runner, not the older external copy
under the e3aa harness directory:

~~~
docs/unified-ops/results/0362g/ppt_phase_abba.py
bytes 12064
sha256 daf58fc8f54ecc30b105f934499ecb313975ebcbfe7a6410356afe6ab699a75c
~~~

It checks the host before, every 100 ms during each child, and after; excludes
only its own process tree and its recognized timer/probe descendants; records
all host rows; requires nonzero phase spans and matching source/target extents;
and verifies each inverse artifact byte-for-byte. Use the existing quiet
coordinator lease and this runner's guard. Do not add a PPT-specific guard.

## Lock and toolchain binding

The final source checkout currently contains:

~~~
Cargo.lock: 139606 bytes
sha256: 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
rust-toolchain.toml sha256: e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
channel: 1.95.0; components: clippy, rustfmt; profile: minimal
~~~

The 0362j reference builds were recorded with lock SHA
40dcbcff741065997d9bc583140a22d921ebe2b4f321570af2587d1cf22d3b5c and the
same 1.95.0 toolchain file. This lock difference is not silently waived:
the candidate build receipt must bind the 4100 lock, registry/package
identities, toolchain file, source tree, probe, and fixture. The protected
reference binaries remain valid historical endpoints; no old c0 lock or source
tree may be copied over the final source.

## Fresh candidate build and correctness handoff

All paths below are external and must be absent before launch. The source is
read-only; only the fresh target, logs, protected candidate binary, and
receipts may be written.

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
TARGET="$ROOT/candidate-target"
BUILD="$ROOT/candidate-build"
TMP="$ROOT/candidate-tmp"
PROTECTED="$BUILD/protected/scoped_patch_probe-472b670547"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"

test "$(git -C "$SOURCE" rev-parse HEAD)" = 472b6705475b073e7fe08f6e11ee17fb1e29f838
test "$(git -C "$SOURCE" rev-parse HEAD^{tree})" = ebc1cee05f124c0c63977aff07a23288c19e7d00
test -z "$(git -C "$SOURCE" status --short)"
test "$(sha256sum "$SOURCE/Cargo.lock" | cut -d' ' -f1)" = 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
test "$(sha256sum "$SOURCE/rust-toolchain.toml" | cut -d' ' -f1)" = e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
test "$(sha256sum "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" | cut -d' ' -f1)" = 0bd4d9c1bcad6319968024ff603ae77195810174e496d5394fbae1838021e549
test "$(sha256sum "$FIXTURE" | cut -d' ' -f1)" = f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a
for path in "$TARGET" "$BUILD" "$TMP"; do test ! -e "$path"; done
mkdir -p "$BUILD" "$TMP"

source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR="$TARGET" TMPDIR="$TMP"
cargo build --offline --locked --release -p litchi-ppt --example scoped_patch_probe \
  >"$BUILD/build.log" 2>&1
install -D -m 0755 "$TARGET/release/examples/scoped_patch_probe" "$PROTECTED"
sha256sum "$PROTECTED" > "$BUILD/candidate-binary.sha256"
~~~

Before native execution, run the same bounded correctness selection recorded
by 0362j against this source and fresh target, with one capped Cargo job:

~~~sh
set -eu
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
TARGET=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c/candidate-target
TMP=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c/candidate-tmp
source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh
export CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_TARGET_DIR="$TARGET" TMPDIR="$TMP"
cd "$SOURCE"
cargo test --offline --locked -p litchi-ppt --all-features --lib scoped::tests -- --test-threads=2
cargo test --offline --locked -p litchi-ppt --all-features --lib -- --test-threads=2
cargo test --offline --locked -p litchi-cfb --all-features --lib -- --test-threads=2
cargo test --offline --locked -p litchi-ppt --all-features --tests -- --test-threads=2
cargo test --offline --locked -p litchi-cfb --all-features --tests -- --test-threads=2
cargo clippy --offline --locked -p litchi-ppt --all-features --all-targets -- -D warnings
cargo clippy --offline --locked -p litchi-cfb --all-features --all-targets -- -D warnings
~~~

Then run the candidate's one-operation preflight outside timing. Its receipt
must bind the candidate binary and fixture and report source_len == 43520, a
positive target length, exit_code == 0, empty stderr, and
inverse_verified == true:

~~~sh
set -eu
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
BUILD="$ROOT/candidate-build"
PROTECTED="$BUILD/protected/scoped_patch_probe-472b670547"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
test -x "$PROTECTED"
test -f "$FIXTURE"
"$PROTECTED" "$FIXTURE" 1 1 1 \
  >"$BUILD/preflight-1-1-1.json" 2>"$BUILD/preflight-1-1-1.stderr"
~~~

## Fresh native ABBA commands

After the candidate build, correctness checks, source/lock receipt and a
quiet-host grant, create a fresh root and run the unchanged 0362g workload
twice. The existing protected references are inputs only. Never use the
historical c0 binary as --after.

~~~sh
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
SOURCE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source
RUNNER="$SOURCE/docs/unified-ops/results/0362g/ppt_phase_abba.py"
FIXTURE="$SOURCE/test-data/unified-ops/0294/video.ppt"
AFTER="$ROOT/candidate-build/protected/scoped_patch_probe-472b670547"
ORIGINAL=/home/zhuhe/code/litchi-worktrees/.coord/ppt-instrumented-baselines-20261021/original-e673d4addb-scoped_patch_probe
SIP=/home/zhuhe/code/litchi-worktrees/.coord/ppt-sip-common-lock-20261021/sip-common-lock-scoped_patch_probe

test ! -e "$ROOT/native-original"; test ! -e "$ROOT/native-sip"
test "$(sha256sum "$RUNNER" | cut -d' ' -f1)" = daf58fc8f54ecc30b105f934499ecb313975ebcbfe7a6410356afe6ab699a75c
mkdir -p "$ROOT"
python3 "$RUNNER" \
  --before "$ORIGINAL" --after "$AFTER" --fixture "$FIXTURE" \
  --out "$ROOT/native-original" \
  --before-commit e673d4addbf4b4224f377e419702a33b3d6958fe \
  --after-commit 472b6705475b073e7fe08f6e11ee17fb1e29f838 \
  --harness "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" \
  --cpu 29 --blocks 5

python3 "$RUNNER" \
  --before "$SIP" --after "$AFTER" --fixture "$FIXTURE" \
  --out "$ROOT/native-sip" \
  --before-commit e4bdab60667470124c4da202477ba6bb1734f256 \
  --after-commit 472b6705475b073e7fe08f6e11ee17fb1e29f838 \
  --harness "$SOURCE/crates/litchi-ppt/examples/scoped_patch_probe.rs" \
  --cpu 29 --blocks 5
~~~

The coordinator must verify that CPU 29 is in the quiet process affinity and
that no unrelated compiler, profiler, probe or benchmark is active. The
runner's manifest, all per-leg host snapshots, output/error/time files and
measurements.json are the native receipt. No ratio or performance claim is
made until both five-block runs are complete and independently checked.

## Allocation handoff

0362j's heap receipts are for c0 and cannot be promoted. The final candidate
needs an allocation/heap receipt at the same full native workload, while the
0362j original and matched-Sip heap receipts remain the historical reference
inputs. Use the reviewed 0362j full-workload protocol, fresh output and an
approved generic process guard; do not run its hard-coded full_heap.py as-is
because it names the c0 source and output directory.

The rebased candidate leg is:

~~~sh
ROOT=/home/zhuhe/code/litchi-worktrees/.coord/ppt-final-source-3431ba566c
ALLOC="$ROOT/candidate-allocation"
BIN="$ROOT/candidate-build/protected/scoped_patch_probe-472b670547"
FIXTURE=/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source/test-data/unified-ops/0294/video.ppt
test ! -e "$ALLOC"
mkdir -p "$ALLOC"
heaptrack --record-only -o "$ALLOC/candidate" "$BIN" "$FIXTURE" 7400 3300 42600
heaptrack_print -f "$ALLOC/candidate.zst" \
  --print-peaks=0 --print-allocators=0 --print-temporary=0 \
  --print-massif "$ALLOC/candidate.massif" >"$ALLOC/candidate.analysis"
~~~

The allocation receipt must record the candidate binary/fixture hashes, exact
counts, host before/after snapshots, heaptrack artifacts, analysis command,
exit code, and the guard's concurrency result. It is allocation/heap evidence
only; it must not be used as timed RSS or latency evidence. The 0362g runner's
per-leg /usr/bin/time -v RSS values remain part of the native receipt.

## Remaining blockers and acceptance boundary

1. The final candidate binary has not been built. No current-source
   correctness, preflight, native, allocation, or performance result exists.
2. The candidate lock is 4100 while both protected 0362j reference builds use
   40dcb. The candidate build must retain that fact and bind registry/package
   and toolchain identities; no lock overlay is permitted.
3. The accepted c0 measurements, c0 heap captures, and primary_noop packet do
   not cover the changed shared.rs, scoped.rs, or slide_order.rs source.
4. A quiet CPU/host grant and independent review of both fresh native receipts
   and the candidate allocation receipt remain required.

This record is preparation only. It makes no native, allocation, fidelity,
performance, or final-gate acceptance claim. The only deferred memo item
remains redaction audit 22 under its existing rules.
