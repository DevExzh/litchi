#!/usr/bin/env bash
# Reproduce change 0582's differential in full.
#
#   run.sh <work-dir>
#
# Builds two extractions of the repository -- committed 93a610ded, and the same
# extraction with change 0580's two files overlaid -- each with its own
# CARGO_TARGET_DIR, regenerates the corpus from this directory's generator, runs
# the harness against both, and classifies every difference.  Nothing is built
# against the shared working tree.
#
# A third, instrumented copy of the after tree is built for the coverage probe.
# That copy is NOT one of the two differential builds: it carries two atomic
# counters added to `strict_layout_for` and to the proof builder, and exists
# only to report how many (input, member, API) triples enter the changed code.

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$HERE/../../../.." && pwd)"
WORK="${1:?usage: run.sh <work-dir>}"
BEFORE_REV=93a610ded
JOBS=8

mkdir -p "$WORK"/{before,after,probe}

# --- the two trees ----------------------------------------------------------
git -C "$REPO" archive --format=tar "$BEFORE_REV" | tar -x -C "$WORK/before"
git -C "$REPO" archive --format=tar "$BEFORE_REV" | tar -x -C "$WORK/after"
# docs/ and media/ are not needed to build soapberry-zip and are several GiB.
rm -rf "$WORK/before/docs" "$WORK/after/docs" "$WORK/before/media" "$WORK/after/media"

# Change 0580's two changed files, taken from the working tree under review.
cp "$REPO/crates/soapberry-zip/src/archive.rs" "$WORK/after/crates/soapberry-zip/src/archive.rs"
cp "$REPO/crates/soapberry-zip/src/office.rs"  "$WORK/after/crates/soapberry-zip/src/office.rs"

for tree in before after; do
  mkdir -p "$WORK/$tree/crates/soapberry-zip/examples"
  cp "$HERE/strict_scope_differential.rs" "$WORK/$tree/crates/soapberry-zip/examples/"
done

# --- corpus -----------------------------------------------------------------
rm -rf "$WORK/corpus"
python3 "$HERE/build_corpus.py" "$REPO" "$WORK/corpus"

# --- build and run, one cargo process at a time -----------------------------
for tree in before after; do
  ( cd "$WORK/$tree" && CARGO_TARGET_DIR="$WORK/target-$tree" \
      cargo build --release -p soapberry-zip --example strict_scope_differential -j "$JOBS" )
done

"$WORK/target-before/release/examples/strict_scope_differential" "$WORK/corpus" "$WORK/report-before.txt"
"$WORK/target-after/release/examples/strict_scope_differential"  "$WORK/corpus" "$WORK/report-after.txt"

python3 "$HERE/classify.py" "$WORK/report-before.txt" "$WORK/report-after.txt" \
    "$WORK/divergences.tsv" > "$WORK/classification-summary.json"

# --- coverage probe (instrumented; not a differential build) ----------------
cp -a "$WORK/after" "$WORK/probe"
rm -f "$WORK/probe/crates/soapberry-zip/examples/strict_scope_differential.rs"
cp "$HERE/strict_scope_coverage.rs" "$WORK/probe/crates/soapberry-zip/examples/"
python3 - "$WORK/probe/crates/soapberry-zip/src/office.rs" <<'PY'
import sys, pathlib
p = pathlib.Path(sys.argv[1]); s = p.read_text()
probe = '''
pub static PROBE_STRICT_TARGETS: std::sync::atomic::AtomicU64 =
    std::sync::atomic::AtomicU64::new(0);
pub static PROBE_STRICT_BUILDS: std::sync::atomic::AtomicU64 =
    std::sync::atomic::AtomicU64::new(0);

'''
anchor = "fn strict_layout_index_error() -> Error {"
assert s.count(anchor) == 1
s = s.replace(anchor, probe + anchor, 1)
old = """    ) -> Result<crate::StrictEntryLayout, Error> {
        strict_layout_for_cached(&self.strict_layout_cache, target, |memo| {
            self.build_strict_layout_proof(target, memo)
        })
    }"""
new = """    ) -> Result<crate::StrictEntryLayout, Error> {
        PROBE_STRICT_TARGETS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        strict_layout_for_cached(&self.strict_layout_cache, target, |memo| {
            PROBE_STRICT_BUILDS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            self.build_strict_layout_proof(target, memo)
        })
    }"""
assert s.count(old) == 2
p.write_text(s.replace(old, new))
PY
( cd "$WORK/probe" && CARGO_TARGET_DIR="$WORK/target-probe" \
    cargo build --release -p soapberry-zip --example strict_scope_coverage -j "$JOBS" )
"$WORK/target-probe/release/examples/strict_scope_coverage" "$WORK/corpus" "$WORK/coverage.tsv"

echo
echo "reports:  $WORK/report-before.txt  $WORK/report-after.txt"
echo "summary:  $WORK/classification-summary.json"
echo "coverage: $WORK/coverage.tsv"
