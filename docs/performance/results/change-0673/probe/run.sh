#!/usr/bin/env bash
# Change 0673's paired ZIP locator-scratch probe.
#
# Usage: run.sh [before-tree] [after-tree]
#
# The probe body and manifest template are reused verbatim from change 0632;
# this packet changes only the two trees under comparison.  Builds and their
# temporary Cargo targets live in a fresh directory outside the repository.
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE="$(cd "$HERE/../../change-0632/probe" && pwd)"
BEFORE_TREE="${1:-/home/zhuhe/code/litchi-worktrees/before-5fa92d7ce}"
AFTER_TREE="${2:-/home/zhuhe/code/litchi-worktrees/0673}"
FIXTURES=(
  "test-data/ooxml/xlsx/ConditionalFormattingSamples.xlsx /xl/worksheets/sheet1.xml xlsx-132"
  "test-data/ooxml/pptx/shapes.pptx /ppt/slides/slide1.xml pptx-shapes"
  "test-data/ooxml/docx/comment.docx /word/document.xml docx-comment"
)

WORK="$(mktemp -d /tmp/litchi-0673-probe.XXXXXX)"
OUT="$HERE/../counts"
mkdir -p "$OUT"

for leg in before after; do
  case "$leg" in
    before) TREE="$BEFORE_TREE" ;;
    after) TREE="$AFTER_TREE" ;;
  esac
  mkdir -p "$WORK/probe-$leg/src"
  cp "$SOURCE/src/main.rs" "$WORK/probe-$leg/src/main.rs"
  sed "s#@TREE@#$TREE#g" "$SOURCE/Cargo.toml.template" > "$WORK/probe-$leg/Cargo.toml"
  echo "== building $leg probe against $TREE"
  (
    cd "$WORK/probe-$leg"
    CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_NET_OFFLINE=true \
      CARGO_TARGET_DIR="$WORK/target-$leg" cargo build --release -j 2
  )
  for spec in "${FIXTURES[@]}"; do
    read -r fixture part tag <<<"$spec"
    echo "== $leg $tag"
    (
      cd "$AFTER_TREE"
      taskset -c 8 "$WORK/target-$leg/release/zip-opc-read-probe" "$fixture" "$part"
    ) > "$OUT/counts-$tag-$leg.txt"
  done
done

sha256sum "$WORK"/target-*/release/zip-opc-read-probe > "$OUT/binaries.sha256"
printf 'before_tree=%s\nafter_tree=%s\nwork_dir=%s\n' \
  "$BEFORE_TREE" "$AFTER_TREE" "$WORK" > "$OUT/provenance.txt"
echo "retained counts in $OUT"
