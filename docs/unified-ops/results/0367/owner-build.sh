#!/usr/bin/env bash
# build.sh SNAPSHOT_DIR OUT_DIR : release-build the three owner libraries, then the shared probe with fat LTO.
set -euo pipefail
src=$(realpath "$1"); out=$(realpath -m "$2"); probe=/home/zhuhe/code/litchi-worktrees/uo2-iwa-w0a/docs/unified-ops/results/0361fd/owner_probe.rs
mkdir -p "$out"
source /deleted-coord-artifact
export CARGO_TARGET_DIR="$out/target" CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_BUILD_JOBS=6
cd "$src"
cargo build --offline --locked --release --lib -p litchi-numbers -p litchi-keynote -p litchi-pages \
  --message-format=json-render-diagnostics > "$out/library.json"
lib() { python3 -c "
import json,sys
for l in open(sys.argv[1]):
  try: i=json.loads(l)
  except: continue
  if i.get('reason')=='compiler-artifact' and i['target']['name']==sys.argv[2]:
    print([f for f in i['filenames'] if f.endswith('.rlib')][0])" "$out/library.json" "$1"; }
rustc --edition=2024 --crate-name owner_probe -C opt-level=3 -C panic=abort -C lto=fat \
  -C embed-bitcode=yes -L "dependency=$CARGO_TARGET_DIR/release/deps" \
  --extern litchi_numbers="$(lib litchi_numbers)" --extern litchi_keynote="$(lib litchi_keynote)" \
  --extern litchi_pages="$(lib litchi_pages)" "$probe" -o "$out/owner-probe" 2>&1 | tee "$out/rustc.log"
rustc -Vv > "$out/rustc-version.txt"
sha256sum "$out/owner-probe" "$probe" "$src/Cargo.lock" > "$out/sha256.txt"
