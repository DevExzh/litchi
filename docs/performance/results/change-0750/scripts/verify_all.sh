#!/bin/bash
# Change 0750: differential campaign, then the test suites of every crate that
# depends on xml-minifier (OOXML, the facade, and read-only checks of the ODF
# and OLE2 crates, which this change does not touch).
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0750
cd /home/zhuhe/code/litchi-worktrees/0750-xml-audit-well-formedness-gaps
OUT=$S/campaign $S/scripts/differential_campaign.sh
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0750 CARGO_BUILD_JOBS=6
timeout 5400 cargo test -p xml-minifier -p litchi-opc -p litchi-ooxml-common -p litchi-docx -p litchi-xlsx -p litchi-pptx -p litchi-xlsb --locked --offline --no-fail-fast > $S/tests-ooxml.log 2>&1
echo "ooxml exit $?" >> $S/verify-status.txt
timeout 5400 cargo test -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --locked --offline --no-fail-fast > $S/tests-facade.log 2>&1
echo "facade exit $?" >> $S/verify-status.txt
timeout 5400 cargo test -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-xls -p litchi-ppt --locked --offline --no-fail-fast > $S/tests-odf-ole.log 2>&1
echo "odf-ole exit $?" >> $S/verify-status.txt
echo done >> $S/verify-status.txt
