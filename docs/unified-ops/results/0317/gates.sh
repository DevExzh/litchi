#!/bin/bash
# Record 0317 gates; one cargo command at a time; each exit status logged.
cd /home/zhuhe/code/litchi-worktrees/uo2-opc-proofs
. /deleted-coord-artifact
LOG=/deleted-coord-artifact
OUT=/deleted-coord-artifact
mkdir -p $OUT; : > $LOG
n=0
run() { n=$((n+1)); "$@" > $OUT/$n.txt 2>&1; echo "$? :: $*" >> $LOG; }
prune() { find target/debug/deps -maxdepth 1 -type f -perm -u+x -newer $OUT/stamp ! -name '*.so' ! -name '*.d' ! -name '*.rlib' ! -name '*.rmeta' -delete 2>/dev/null; touch $OUT/stamp; }
FEATS=doc,docx,ppt,pptx,xls,xlsx,xlsb,rtf,odt,ods,odp,oth,odm,odg,odc,odi,odb,odf-formula,odf-objects,pages,keynote,numbers,iwork-adapter,opc,markdown,formula,yaml,eval,test-support
touch $OUT/stamp
run cargo fmt --all --check
run cargo clippy -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-pptx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings
for f in "" docx pptx; do
  if [ -z "$f" ]; then run cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings
  else run cargo clippy -p litchi-ops --no-default-features --features $f --all-targets -- -D warnings; fi
done
for c in xml-minifier litchi-opc litchi-docx litchi-pptx; do run cargo test -p $c --all-features; prune; done
run cargo test -p litchi-ops --all-features --lib; prune
run cargo test -p litchi-ops --all-features --doc; prune
tests=$(ls crates/litchi-ops/tests/*.rs | xargs -n1 basename | sed 's/\.rs$//')
set -- $tests
while [ $# -gt 0 ]; do
  args=""; i=0
  while [ $# -gt 0 ] && [ $i -lt 20 ]; do args="$args --test $1"; shift; i=$((i+1)); done
  run cargo test -p litchi-ops --all-features $args; prune
done
run cargo test -p litchi-ops --no-default-features --features $FEATS --test capabilities --test registry; prune
run cargo test -p litchi --all-features; prune
run cargo test -p litchi-cli --all-features --test officecli_corpus; prune
RUSTDOCFLAGS=-Dwarnings run cargo doc --no-deps --all-features -p xml-minifier -p litchi-opc -p litchi-docx -p litchi-pptx
run python3 tools/check_crate_boundaries.py
run python3 tools/non_iwork_gate.py verify
run python3 tools/test_no_3rdparty_references.py
RUSTC_BOOTSTRAP=1 run python3 tools/check_api_glossary.py --target-dir target/glossary
run crates/pyo3-litchi/tests/run_tests.sh
run bash -c 'cd tools/perf-baseline && cargo check --locked --all-targets'
echo DONE >> $LOG
