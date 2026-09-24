#!/bin/bash
# Record 0764 ABBA campaign: run from the worktree root (the benign cases read
# fixtures by relative path). Pinned to core 16.
set -u
S=/home/zhuhe/code/litchi-worktrees/scratch/0764
OUT=$S/abba
A=$S/bin/a
B=$S/bin/b
mkdir -p $OUT
for c in mce_declaration_flood mce_shadowed_chain mce_hoisting mce_declarations_admitted mce_stream_declarations docx_styles_duplicates audit_attribute_flood; do
  python3 $S/abba.py --kind bounds --a $A/xml_attribute_bounds --b $B/xml_attribute_bounds --case $c --rounds 4 --samples 15 --warmup 3 --core 16 --out $OUT/$c >> $OUT/campaign.log 2>&1 || echo "FAILED $c" >> $OUT/campaign.log
done
for c in mce_benign_worksheet mce_benign_document mce_stream_benign_worksheet audit_benign_worksheet docx_styles_benign; do
  python3 $S/abba.py --kind bounds --a $A/xml_attribute_bounds --b $B/xml_attribute_bounds --case $c --rounds 4 --samples 31 --warmup 3 --core 16 --out $OUT/$c >> $OUT/campaign.log 2>&1 || echo "FAILED $c" >> $OUT/campaign.log
done
for c in docx_semantic_full_text xlsx_first_cell pptx_semantic_full_text xlsx_source_backed_cell_values_one_edit_save docx_source_backed_one_edit_save; do
  python3 $S/abba.py --kind harness --a $A/litchi-perf-baseline --b $B/litchi-perf-baseline --case $c --rounds 4 --samples 15 --warmup 3 --core 16 --out $OUT/$c >> $OUT/campaign.log 2>&1 || echo "FAILED $c" >> $OUT/campaign.log
done
echo done > $OUT/campaign.status
