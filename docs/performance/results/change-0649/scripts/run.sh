#!/usr/bin/env bash
# Change 0649: reproduce the attribution of change 0638's 133.61 ms
# `pptx_real_file_ordinary_save_edit` row.
#
# Usage: run.sh <checkout> <probe-dir> <out> <cpu>
#
# <checkout> is a checkout of the base commit; <probe-dir> is a scratch
# directory outside the repository holding `probe/Cargo.toml.example` renamed to
# `Cargo.toml` with `<checkout>` substituted, plus `probe/src/main.rs`.
#
# Three legs, in the order the record reports them:
#   1. deterministic counts, on the INSTRUMENTED leg (apply instrumentation.patch
#      and copy perf0649.rs into crates/litchi-ooxml-common/src/ first);
#   2. native cycles, instructions and per-symbol shares, on the CLEAN leg;
#   3. paired wall clock with three repeats, on the CLEAN leg.
set -euo pipefail
CHECKOUT=${1:?checkout}
PROBE=${2:?probe directory}
OUT=${3:?output directory}
CPU=${4:?cpu}
REAL="${CHECKOUT}/test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
GEN="generated:12x8"   # the shape of build_semantic_pptx_corpus(Medium): 12 slides x 8 text boxes
mkdir -p "$OUT"

# Stage the binary outside the Cargo target directory (0627's lesson).
CARGO_TARGET_DIR="${OUT}/target" cargo build --release --manifest-path "${PROBE}/Cargo.toml"
cp "${OUT}/target/release/probe0649" "${OUT}/probe0649"
BIN="${OUT}/probe0649"

# The marker-stripped control: every occurrence of the MCE namespace URI is
# replaced by an equal-length URI the codec does not recognize, so the archive
# keeps its member count, its element count and its uncompressed byte total and
# differs only in which branch of `process_markup_compatibility` it takes.
python3 - "$REAL" "${OUT}/marker-stripped.pptx" <<'PY'
import sys, zipfile
OLD = b'http://schemas.openxmlformats.org/markup-compatibility/2006'
NEW = b'http://schemas.openxmlformats.org/markup-kompatibility/2006'
assert len(OLD) == len(NEW)
source, destination = sys.argv[1], sys.argv[2]
zi = zipfile.ZipFile(source)
zo = zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED)
for info in zi.infolist():
    zo.writestr(info, zi.read(info.filename).replace(OLD, NEW))
zo.close()
PY

# 1. Deterministic counts (instrumented leg only; `counts` needs --features counts).
for source in "$REAL" "$GEN" "${OUT}/marker-stripped.pptx"; do
  taskset -c "$CPU" "$BIN" counts "$source" > "${OUT}/counts-$(basename "$source").tsv" || true
done
taskset -c "$CPU" "$BIN" allocations "$REAL" > "${OUT}/allocations-real.tsv"
taskset -c "$CPU" "$BIN" allocations "$GEN"  > "${OUT}/allocations-generated.tsv"

# 2. Native cycles and instructions: isolation pairs over the six prefix stages.
for stage in open capture transaction settext commit apply; do
  for n in 2 10; do
    echo "== ${stage} n=${n}"
    taskset -c "$CPU" perf stat -x, -e cycles,instructions,task-clock \
      -- "$BIN" prefix "$REAL" "$stage" "$n" 2>&1 | grep -E "^[0-9.]+,,"
  done
done > "${OUT}/perfstat-isolation-pairs.txt"

taskset -c "$CPU" perf record -F 999 --call-graph=dwarf,4096 -o "${OUT}/capture.data" \
  -- "$BIN" prefix "$REAL" capture 40
taskset -c "$CPU" perf record -F 999 -g -o "${OUT}/edit.data" \
  -- "$BIN" prefix "$REAL" apply 20
perf report -i "${OUT}/capture.data" --no-children -g none --stdio > "${OUT}/perf-capture.txt"
perf report -i "${OUT}/edit.data"    --no-children -g none --stdio > "${OUT}/perf-edit.txt"

# The callgrind isolation pair. `prefix ... capture` derives no edit target, so
# the profile is the captures and nothing else.
for n in 2 6; do
  taskset -c "$CPU" valgrind --tool=callgrind --cache-sim=no --branch-sim=no \
    --callgrind-out-file="${OUT}/cg-capture-${n}.out" "$BIN" prefix "$REAL" capture "$n"
done
callgrind_annotate --inclusive=yes --threshold=100 "${OUT}/cg-capture-6.out" > "${OUT}/callgrind-capture.txt"

# 3. Paired wall clock, three repeats, for the A/A floor.
for repeat in R1 R2 R3; do
  for label_source in "real:${REAL}" "generated:${GEN}" "control:${OUT}/marker-stripped.pptx"; do
    label=${label_source%%:*}; source=${label_source#*:}
    taskset -c "$CPU" "$BIN" phases "$source" 30 > "${OUT}/phases-${label}-${repeat}.tsv"
  done
done
echo "complete"
