#!/bin/bash
cd /home/zhuhe/code/litchi-worktrees/base-009d515bef
B=/home/zhuhe/code/litchi-worktrees/targets/base-009d515bef/release/litchi-perf-baseline
O=/tmp/claude-1001/-home-zhuhe-code-litchi/e8f8ff4d-7cfa-4135-8755-9c3e3f5fc16d/scratchpad/sweep
taskset -c 20 $B --samples 7 --warmup 2 --semantic-shape medium,large --case docx_semantic_open,docx_semantic_full_text,docx_semantic_one_edit_save,docx_semantic_noop_edit_save,docx_semantic_one_percent_edit_save,docx_semantic_create_small,docx_streaming_create,pptx_semantic_open,pptx_semantic_full_text,pptx_semantic_one_edit_save,pptx_semantic_noop_edit_save,pptx_semantic_one_percent_edit_save,pptx_streaming_create,doc_semantic_open,doc_semantic_full_text,doc_semantic_one_edit_save,doc_semantic_noop_edit_save,xls_semantic_open,xls_semantic_full_cell_scan,xls_semantic_one_edit_save,xls_semantic_noop_edit_save,ppt_semantic_open,ppt_semantic_full_text,ppt_semantic_one_edit_save,ppt_semantic_noop_edit_save,doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to --json $O/semantic.json > $O/semantic.log 2>&1
echo semantic-done $? >> $O/status
taskset -c 20 $B --samples 7 --warmup 2 --xlsx-shape medium,dense-wide --case xlsx_open_owned,xlsx_full_cell_scan,xlsx_first_cell,xlsx_noop_commit_save,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save,xlsx_streaming_create --json $O/xlsx.json > $O/xlsx.log 2>&1
echo xlsx-done $? >> $O/status
taskset -c 20 $B --samples 7 --warmup 2 --case docx_ordinary_save_lifecycle,xlsx_ordinary_save_lifecycle,pptx_ordinary_save_lifecycle,pptx_cross_copy_plain_lifecycle,pptx_cross_copy_media_rich_lifecycle,pptx_source_backed_cross_copy_media_rich_lifecycle,pptx_source_backed_one_edit_save,docx_source_backed_one_edit_save,xlsx_source_backed_cell_values_one_edit_save,xlsx_eager_cell_values_one_edit_save,xlsx_producer_dense_source_one_edit_save,xlsx_producer_medium_source_one_edit_save,opc_noop_save,opc_mutated_save --json $O/misc.json > $O/misc.log 2>&1
echo misc-done $? >> $O/status
