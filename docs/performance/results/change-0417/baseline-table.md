# Normal repeat observations

All elapsed values are milliseconds, shown as repeat 1 / repeat 2.
Read `timing-boundaries.md` for each selector. These intervals differ;
cross-selector latency ratios do not represent speedups. The >5% flags
use absolute signed repeat drift and preserve both faster and slower tails.

| Selector | p50 ms | p95 ms | p99 ms | >5% drift |
| --- | ---: | ---: | ---: | --- |
| `cfb_list_streams` | 0.009210 / 0.009120 | 0.009380 / 0.009280 | 0.012320 / 0.012190 | none |
| `cfb_read_one` | 0.000540 / 0.000540 | 0.000580 / 0.000580 | 0.000600 / 0.000590 | none |
| `doc_fresh_write_to` | 0.185315 / 0.183746 | 0.200680 / 0.200041 | 0.209641 / 0.207611 | none |
| `docx_story_hyperlink_redaction_save` | 0.441852 / 0.444262 | 0.452412 / 0.454772 | 0.458492 / 0.459652 | none |
| `odf_mimetype_repair_plan` | 0.555972 / 0.555678 | 0.561813 / 0.561832 | 0.565252 / 0.566203 | none |
| `odf_validation_report` | 0.034330 / 0.033900 | 0.035361 / 0.034840 | 0.043940 / 0.041010 | p99 |
| `odp_semantic_text_to_sink` | 0.029495 / 0.029290 | 0.029800 / 0.029630 | 0.036150 / 0.038270 | p99 |
| `ods_semantic_text_to_sink` | 0.118121 / 0.114910 | 0.125011 / 0.120860 | 0.126351 / 0.123051 | none |
| `odt_semantic_text_to_sink` | 0.039760 / 0.039680 | 0.042670 / 0.042570 | 0.048710 / 0.049250 | none |
| `opc_open` | 0.653143 / 0.653652 | 0.659873 / 0.667293 | 0.663302 / 0.670783 | none |
| `ppt_fresh_write_to` | 0.138266 / 0.138426 | 0.153020 / 0.152951 | 0.157931 / 0.160391 | none |
| `pptx_cross_copy_media_rich` | 1151.843381 / 1152.214207 | 1154.190096 / 1154.143199 | 1155.262662 / 1155.101118 | none |
| `pptx_source_backed_cross_copy_plain` | 1.739007 / 1.739922 | 1.750858 / 1.752088 | 1.757917 / 1.759537 | none |
| `rtf_semantic_split_paragraph_save` | 0.149675 / 0.151940 | 0.157510 / 0.159301 | 0.161491 / 0.162311 | none |
| `rtf_semantic_text_to_sink` | 0.012300 / 0.012330 | 0.012450 / 0.012490 | 0.017990 / 0.018430 | none |
| `rtf_streaming_create` | 8.043769 / 8.425581 | 8.058615 / 8.451797 | 8.069224 / 8.467456 | none |
| `xls_fresh_write_to` | 1.205910 / 1.244370 | 1.216625 / 1.253776 | 1.223535 / 1.256166 | none |
| `xls_validation_report` | 0.863383 / 0.963114 | 0.866293 / 0.966164 | 0.868524 / 0.969084 | p50, p95, p99 |
| `xlsx_eager_cell_remove_edit_save` | 9.413110 / 9.461485 | 9.459150 / 9.526651 | 9.527930 / 9.552121 | none |
| `xlsx_eager_defined_names_edit_save` | 235.950286 / 235.608483 | 236.737071 / 238.265780 | 241.453646 / 449.867652 | p99 |
| `xlsx_eager_merge_commit_save` | 0.050390 / 0.050211 | 0.059150 / 0.058810 | 0.061971 / 0.062270 | none |
| `xlsx_eager_row_visibility_edit_save` | 27.249479 / 27.333223 | 27.391567 / 27.521729 | 27.454257 / 27.608199 | none |
| `xlsx_eager_sheet_protection_edit_save` | 235.727695 / 235.638229 | 236.531209 / 236.431650 | 238.470862 / 245.488636 | none |
| `xlsx_full_cell_scan` | 0.606732 / 0.609023 | 0.611422 / 0.613462 | 0.613433 / 0.615633 | none |
| `xlsx_join_disjoint_commit_save` | 11.636139 / 11.650095 | 11.700259 / 11.759270 | 11.760529 / 11.799430 | none |
| `xlsx_list_sheets` | 0.000110 / 0.000110 | 0.000120 / 0.000120 | 0.000220 / 0.000240 | p99 |
| `xlsx_one_cell_commit` | 2.398835 / 2.368300 | 2.414520 / 2.386411 | 2.425920 / 2.396141 | none |
| `xlsx_one_percent_commit` | 9.611944 / 9.573821 | 9.662819 / 9.619762 | 9.716290 / 9.633741 | none |
| `xlsx_streaming_create` | 12.697182 / 12.804115 | 12.751072 / 12.913416 | 12.775872 / 12.937706 | none |
| `xlsx_three_way_disjoint_commit_save` | 11.646579 / 11.680829 | 11.735089 / 11.758149 | 11.779179 / 11.795089 | none |

IID median order-statistic intervals and full vectors remain in
`summary.json` and the raw reports. These intervals do not model
shared-host drift or establish independence of successive samples.

## Throughput and whole-process RSS

Throughput is the reciprocal of each scoped p50 (operations/second).
It is not physical input bandwidth. RSS is `/usr/bin/time -v` process
maximum, including corpus construction, setup, warmups, verification
and retained output, and is not an operation-local memory peak.

| Selector | Operations/s at p50, repeat 1 / 2 | Peak RSS KiB, repeat 1 / 2 |
| --- | ---: | ---: |
| `cfb_list_streams` | 108577.633 / 109649.123 | 74552 / 74660 |
| `cfb_read_one` | 1851851.852 / 1851851.852 | 74676 / 74628 |
| `doc_fresh_write_to` | 5396.217 / 5442.295 | 74612 / 74676 |
| `docx_story_hyperlink_redaction_save` | 2263.201 / 2250.924 | 74684 / 74628 |
| `odf_mimetype_repair_plan` | 1798.652 / 1799.603 | 74636 / 74672 |
| `odf_validation_report` | 29129.042 / 29498.525 | 74628 / 74656 |
| `odp_semantic_text_to_sink` | 33904.052 / 34141.345 | 74660 / 74672 |
| `ods_semantic_text_to_sink` | 8465.895 / 8702.463 | 74664 / 74548 |
| `odt_semantic_text_to_sink` | 25150.905 / 25201.613 | 74640 / 74664 |
| `opc_open` | 1531.058 / 1529.866 | 74672 / 74536 |
| `ppt_fresh_write_to` | 7232.436 / 7224.076 | 74668 / 74684 |
| `pptx_cross_copy_media_rich` | 0.868 / 0.868 | 820188 / 820556 |
| `pptx_source_backed_cross_copy_plain` | 575.041 / 574.738 | 74676 / 74676 |
| `rtf_semantic_split_paragraph_save` | 6681.142 / 6581.545 | 74676 / 74676 |
| `rtf_semantic_text_to_sink` | 81300.813 / 81103.001 | 74592 / 74484 |
| `rtf_streaming_create` | 124.320 / 118.686 | 74612 / 74668 |
| `xls_fresh_write_to` | 829.249 / 803.620 | 74668 / 74612 |
| `xls_validation_report` | 1158.235 / 1038.299 | 74676 / 74584 |
| `xlsx_eager_cell_remove_edit_save` | 106.235 / 105.692 | 74676 / 74644 |
| `xlsx_eager_defined_names_edit_save` | 4.238 / 4.244 | 131764 / 131712 |
| `xlsx_eager_merge_commit_save` | 19845.207 / 19915.955 | 74676 / 74676 |
| `xlsx_eager_row_visibility_edit_save` | 36.698 / 36.586 | 74604 / 74624 |
| `xlsx_eager_sheet_protection_edit_save` | 4.242 / 4.244 | 131644 / 131648 |
| `xlsx_full_cell_scan` | 1648.174 / 1641.974 | 74676 / 74676 |
| `xlsx_join_disjoint_commit_save` | 85.939 / 85.836 | 122240 / 118688 |
| `xlsx_list_sheets` | 9090909.091 / 9090909.091 | 74612 / 74624 |
| `xlsx_one_cell_commit` | 416.869 / 422.244 | 74612 / 74536 |
| `xlsx_one_percent_commit` | 104.037 / 104.452 | 74552 / 74612 |
| `xlsx_streaming_create` | 78.758 / 78.100 | 74564 / 74612 |
| `xlsx_three_way_disjoint_commit_save` | 85.862 / 85.610 | 118908 / 122496 |

## Available allocation attribution

Both repeats have identical p50/p95/p99 allocation calls and bytes
for the two attributed selectors. The other 28 remain unavailable.

| Selector | Allocation calls | Allocated bytes |
| --- | ---: | ---: |
| `docx_story_hyperlink_redaction_save` | 4518 | 8548090 |
| `rtf_streaming_create` | 16387 | 1507536 |
