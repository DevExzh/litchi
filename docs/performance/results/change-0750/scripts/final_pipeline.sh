#!/bin/bash
# Change 0750: on the final candidate, rebuild the after-leg harness with the
# command the before leg used, then run the differential campaign and every
# dependent crate's tests.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps
( CARGO_BUILD_JOBS=6 CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 \
    cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline ) > $S/build-after-final.log 2>&1
echo "after-final exit $?" >> $S/build-legs-status.txt
cp /home/zhuhe/code/litchi-worktrees/targets/0750/release/litchi-perf-baseline $S/bin/litchi-perf-baseline.after
sha256sum $S/bin/litchi-perf-baseline.* >> $S/build-legs-status.txt
rm -rf $S/campaign; OUT=$S/campaign $S/scripts/differential_campaign.sh
echo "campaign done" >> $S/verify-status.txt
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 CARGO_BUILD_JOBS=6
timeout 5400 cargo test -p xml-minifier -p litchi-opc -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-xlsb --locked --offline --no-fail-fast > $S/tests-ooxml.log 2>&1
echo "ooxml exit $?" >> $S/verify-status.txt
timeout 5400 cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline --no-fail-fast > $S/tests-facade.log 2>&1
echo "facade exit $?" >> $S/verify-status.txt
timeout 5400 cargo test -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-xls -p litchi-ppt --locked --offline --no-fail-fast > $S/tests-odf-ole.log 2>&1
echo "odf-ole exit $?" >> $S/verify-status.txt
echo done >> $S/verify-status.txt
