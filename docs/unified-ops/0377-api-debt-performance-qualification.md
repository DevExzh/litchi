# 0377 — Performance qualification of the API-debt program

- Status: implemented on `uo2/fin-perf`, cut from `feat/unified-ops-api` at
  `ac0d111c4c` (the commit on which all 53 gate rows passed, record 0373),
  trunk merged at `41dc20503e` (trunk tip `c4e4c5266a`).
- Date: 2026-10-26
- Input: owner decision [0350](0350-owner-api-debt-decision.md), contract
  [0351](0351-api-contract.md) §2.6 and §5, `docs/UNIFIED_OPS_GOAL.md`
  (PERFORMANCE BAR), owner decision 2026-10-26 (ODT/ODP performance
  improvement deferred).
- Baseline: trunk before the API-debt program, `839212fa9d` (record 0349).
  It builds; no substitute ancestor was needed.
- Candidate: this branch after the trunk merge, `41dc20503e` plus record
  commits (the fixes of §3 on top of `ac0d111c4c`).
- `performance_claim`: measured (§4–§6). Callgrind instruction counts of the
  timed region are the primary evidence for the 5 % regression rule;
  wall-clock A-B-B-A blocks judge the leaf-vs-unified budgets.

## 1. Summary

| Item | Result |
|---|---|
| Routes measured | 98 paired leaf/unified routes of the 21 claimed formats (`unified-fin-perf`), 41 perf-baseline default cases (213 rows), 14 `odf_w6_probe` routes, 5 source-backed ODF lifecycles, 7 Python-binding routes |
| Regressions above 5 % (instructions) found on `ac0d111c4c` and fixed here | PPTX edits +8–15 %, PPT single text set (unified 1.41× leaf), DOC/PPT `to_bytes` (1.68× / 1.83× leaf), XLSX 1k batch +33 %, XLSX single set +6.7 %, XLSB removals +6.8 %, ODC open+get +16 % / set +10 % |
| Remaining above 5 % (instructions), outside the ODT/ODP deferral | OOXML metadata commit (DOCX 1.66, XLSX 1.17, PPTX 1.06: fixed ≈1–3 µs, §5.2), XLSX unified cell set (single 1.052, 1k batch 1.095: the contract snapshot bound, §5.3), ODC open+get on the odfdom fixture (1.080, raw-XML admission, §5.4) |
| Deferred by owner (measured, not fixed) | ODT open+get 1.50, ODT set 1.18, flat ODT open+get 1.23 (`odf_w6_probe`, instructions 1.49 / 1.16 / 1.53); ODT-engine routes of OTH/ODM; ODP existing-append lifecycle 1.04 Ir / 1.06 wall. ODP memory residual: **closed** — tail-append peak heap 0.58 MiB on both builds (1.000), outputs identical |
| Leaf-vs-unified budgets (wall, default release) | all met except XLSX single cell set 1.108 (1.095 with `codegen-units=1`, 1.081 in instructions) and ODC open (+70 µs, 1.057; pre-existing 1.071 at baseline) |
| 1k-op batch (2.0×) | met by every format; worst XLSX 1.76, ODP 1.67, ODM 1.65 |
| Second save no slower than the old first save | met everywhere (§4.8) |
| Wall-clock vs instruction disagreements | perf-baseline OPC opens and XLS write +7–10 % wall at equal instructions; rebuilt with `codegen-units=1` both builds measure 0.995–1.011 (§5.5) |
| Broken existing scenario found | perf-baseline default matrix failed on trunk since record 0358 (stale-source oracle); fixed (§3) |

## 2. Method

- **Harness** `tools/perf-baseline/unified-fin-perf` (new; standalone
  workspace like `unified-batch-closure`). One source builds on the baseline
  (`--features base`, two compat shims: `batch_with_limits` and the OLE
  `properties_mut` signature) and on the candidate. It compiles each tree's
  own `unified-batch-closure` format modules (`#[path]`), so the record 0337
  scenarios and their leaf calls are that tree's. Groups:
  - `batch`: one atomic `File::batch` of 1,000 items against one hand-written
    leaf edit, fixture of 1,000 targets (ODI: its six frame properties);
  - `set`: the same case at one operation, the unified leg calling
    `File::set`/`File::remove` with the item's path and properties;
  - `extra`: open (default and eager unified owner; DOCX eager paired with
    the leaf open plus the eager owner's validation text pass, as the
    perf-baseline `docx_file_leaf_eager_*` legs do), first and second save
    of a freshly opened owner after a metadata edit, metadata title set,
    DOCX paragraph add, DOC paragraph get, PPTX slide remove/move and shape
    remove.
  Every timed call runs inside the non-inlined `fin_perf_region`.
- **Instructions**: `valgrind --tool=callgrind --collect-atstart=no
  --toggle-collect='*fin_perf_region*'`, one process per route and leg, three
  calls after an uncounted warm-up (one call for batches); Ir per call
  (`results/0377/callgrind.tsv`, scripts `callgrind-one.sh`,
  `callgrind-list.sh`). `odf_w6_probe` and perf-baseline cases have no
  region marker: their cost per round/sample is the difference of two runs
  (6 − 2 rounds; 9 − 3 samples), which cancels setup.
- **Wall clock**: `taskset -c 47`, binary order base, candidate, candidate,
  base; inside each process every route is a leaf/unified ABBA of 5 blocks ×
  10 rounds (`abba.sh`; `abba.json` keeps every block median). The four
  routes whose leaf costs 40–170 G instructions per batch (DOC, ODG, ODB,
  ODF formula 1k batches) have instructions only. Host load average 15–30
  (other agents' builds).
- **Builds**: release, the workspace profile (fat LTO, `panic = "abort"`;
  harness `debug = 1`), `CARGO_BUILD_JOBS=6`, separate targets
  (`build.sh`). Binary SHA-256 and sizes: `binaries.sha256`; binaries are
  kept outside the repository.

## 3. Fixes

| Commit | Change | Effect (instructions, candidate / baseline) |
|---|---|---|
| `937bc427e7` | PPTX: the package memo carries the capture's slide-root memo under its source profile; every detached `Package::edit` reused only digests and rescanned each slide's physical limits and notes root. The slide-removal plan capture reuses its parent snapshot's memos. One test renamed for the new retention (one more `Arc` per slide payload, no extra bytes pinned) | shape text set 1.149 → 0.934; batch 1.148 → 0.934; slide move/remove, shape remove 1.08–1.11 → 0.73–0.91 |
| `38b646a062`, `de13cd37fd`, `b370ff1699` | XLSX: commit read only the capture ledger's maximum after normalization (`check_patch` measures the patch itself) but rebuilt the exact ledger twice; joins still rebuild it. The coarse cell reservation is skipped once the worksheet before-image is admitted (the setter charges the cell). The unified run's staged-target set is folded into its charge map | 1k unified batch 1.330 → 1.095; single set 1.067 → 1.052 |
| `2db4736472` | XLSB: `ensure_prepared` returned after cloning the worksheet part name on every staged operation | removals 1.068 → 1.014 |
| `468bce40b9` | ops: a single PPT shape-text `File::set` publishes through the one-item coalesced run (the native chunked text edit) and falls back to the ordinary writer for anything else or on any error; `File::to_bytes` of DOC/PPT uses the leaf's single sized copy instead of the sink route's chunk-by-chunk double copy | PPT set unified/leaf 1.41 → 1.005; DOC/PPT save 1.68/1.83 → 1.00 |
| `92f2b5e4ee` | ODF: namespace URI admission walked each declaration character by character; a printable-ASCII value without `&` normalizes to itself | ODC open+get 1.162 → 1.080, set 1.096 → 1.048 |
| `b370ff1699` | perf-baseline: the ODP existing-append oracle expected the pre-0358 `InvalidFormat` stale-source message; since record 0358 a stale patch source is a typed `PatchConflict`, so the default matrix stopped at `odp_existing_append_lifecycle` on trunk | default matrix runs again |

Every fix keeps outputs byte-identical (the harness compares every leaf and
unified output outside the clock: `exact_bytes`/`decoded_members` per
route); `litchi-pptx`, `litchi-xlsx`, `litchi-xlsb`, `litchi-odf-common`,
`litchi-odc` and the litchi-ops DOC/PPT/XLSX test targets pass (§7).

## 4. Per-wave tables

Columns: unified-leg instructions per call on the baseline and the
candidate (ratio = the 5 % regression rule), the leaf leg's ratio, the
candidate's wall-clock medians and unified/leaf ratio with the spread of the
20 block ratios, the budget and the verdict ("meets"/"miss" for the budget;
"> 5 % Ir" when either leg regressed above 5 % in instructions).
Generated by `results/0377/table.py` from the committed evidence.

### W1 DOCX

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `docx_default_open` | 367,620 → 374,772 (1.019) | 1.013 | 16.8 / 11.1 | 0.663 [0.658–0.667] | 1.05 or +20 µs | meets |
| `docx_eager_run_bold_set_single` | 5,010,275 → 5,005,467 (0.999) | 0.998 | 181.6 / 188.8 | 1.039 [1.034–1.049] | 1.10 | meets |
| `docx_eager_open_validated` | 5,593,265 → 5,574,613 (0.997) | 0.997 | 222.4 / 222.4 | 1.000 [0.994–1.004] | 1.05 or +20 µs | meets |
| `docx_eager_metadata_title_set` | 97,224 → 161,630 (1.662) | 1.662 | 7.7 / 7.9 | 1.033 [1.027–1.044] | 1.10 | meets, >5 % Ir |
| `docx_eager_run_bold_set_batch1000` | 456,565,765 → 444,698,127 (0.974) | 0.989 | 7,381.6 / 10,444.1 | 1.418 [1.409–1.422] | 2.0 | meets |
| `docx_eager_save_first` | 3,164,582 → 3,194,796 (1.010) | 1.009 | 40.2 / 41.0 | 1.021 [1.015–1.026] | 1.10 | meets |
| `docx_eager_save_second` | 3,166,313 → 3,190,450 (1.008) | 1.007 | 36.1 / 36.1 | 1.002 [0.999–1.008] | 1.10 | meets |
| `docx_eager_paragraph_add` | 5,592,849 → 5,595,366 (1.000) | 1.000 | 521.5 / 241.3 | 0.463 [0.456–0.471] | 1.10 | meets |

### W2 PPTX

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `pptx_default_open` | 3,695,094 → 3,800,753 (1.029) | 1.012 | 307.1 / 157.6 | 0.507 [0.499–0.521] | 1.05 or +20 µs | meets |
| `pptx_eager_shape_text_set_batch1000` | 460,808,104 → 430,127,697 (0.933) | 0.936 | 9,023.0 / 9,484.4 | 1.050 [1.050–1.054] | 2.0 | meets |
| `pptx_eager_open` | 6,168,146 → 6,321,730 (1.025) | 1.010 | 311.0 / 232.7 | 0.750 [0.738–0.756] | 1.05 or +20 µs | meets |
| `pptx_eager_shape_text_set_single` | 292,363,318 → 273,010,649 (0.934) | 0.934 | 8,796.5 / 8,829.3 | 1.002 [0.997–1.006] | 1.10 | meets |
| `pptx_eager_metadata_title_set` | 1,899,888 → 2,014,949 (1.061) | 1.061 | 75.0 / 75.9 | 1.009 [1.004–1.022] | 1.10 | meets, >5 % Ir |
| `pptx_eager_save_first` | 2,401,152 → 2,420,768 (1.008) | 1.000 | 1,251.4 / 86.4 | 0.069 [0.068–0.073] | 1.10 | meets |
| `pptx_eager_save_second` | 2,466,004 → 2,473,349 (1.003) | 0.997 | 1,223.3 / 69.6 | 0.057 [0.056–0.058] | 1.10 | meets |
| `pptx_eager_slide_move` | 19,517,223 → 14,202,714 (0.728) | 0.733 | 518.6 / 517.0 | 0.997 [0.990–1.003] | 1.10 | meets |
| `pptx_eager_slide_remove` | 230,519,018 → 210,705,926 (0.914) | 0.915 | 6,490.5 / 6,500.1 | 1.001 [0.997–1.004] | 1.10 | meets |
| `pptx_eager_shape_remove` | 21,912,189 → 16,391,885 (0.748) | 0.752 | 608.8 / 609.1 | 0.998 [0.994–1.005] | 1.10 | meets |

### W3 PPT

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `ppt_eager_shape_text_set_batch1000` | 105,242,903 → 98,255,977 (0.934) | 0.929 | 2,360.8 / 2,697.2 | 1.143 [1.131–1.151] | 2.0 | meets |
| `ppt_eager_shape_text_set_single` | 19,758,005 → 13,855,396 (0.701) | 0.986 | 278.5 / 283.2 | 1.017 [1.005–1.025] | 1.10 | meets |
| `ppt_default_open` | 628,085 → 627,493 (0.999) | 0.998 | 11.6 / 11.8 | 1.009 [1.007–1.014] | 1.05 or +20 µs | meets |
| `ppt_eager_open` | 627,972 → 626,896 (0.998) | 0.997 | 11.7 / 11.9 | 1.012 [1.007–1.023] | 1.05 or +20 µs | meets |
| `ppt_eager_metadata_title_set` | 3,931,978 → 3,875,473 (0.986) | 0.985 | 70.0 / 72.2 | 1.031 [1.026–1.037] | 1.10 | meets |
| `ppt_eager_save_first` | 753,803 → 423,420 (0.562) | 1.000 | 2.6 / 2.7 | 1.029 [1.002–1.056] | 1.10 | meets |
| `ppt_eager_save_second` | 774,536 → 423,225 (0.546) | 1.000 | 1.9 / 1.9 | 1.005 [0.997–1.016] | 1.10 | meets |

### W4 XLSX/XLSB

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `xlsx_eager_cell_value_set_batch1000` | 60,094,186 → 65,776,858 (1.095) | 1.062 | 1,170.7 / 2,063.9 | 1.764 [1.753–1.768] | 2.0 | meets, >5 % Ir |
| `xlsx_eager_cell_value_set_single` | 460,257 → 484,062 (1.052) | 1.007 | 20.0 / 22.2 | 1.108 [1.099–1.116] | 1.10 | miss, >5 % Ir |
| `xlsx_default_open` | 809,488 → 829,381 (1.025) | 1.018 | 32.0 / 24.8 | 0.776 [0.769–0.781] | 1.05 or +20 µs | meets |
| `xlsx_eager_open` | 1,117,504 → 1,138,424 (1.019) | 1.018 | 31.8 / 32.2 | 1.014 [1.008–1.017] | 1.05 or +20 µs | meets |
| `xlsx_eager_metadata_title_set` | 185,152 → 215,879 (1.166) | 1.113 | 9.0 / 9.9 | 1.096 [1.091–1.112] | 1.10 | meets, >5 % Ir |
| `xlsx_eager_save_first` | 3,345,476 → 3,370,941 (1.008) | 1.008 | 46.8 / 47.5 | 1.016 [1.004–1.024] | 1.10 | meets |
| `xlsx_eager_save_second` | 3,344,322 → 3,369,087 (1.007) | 1.007 | 42.3 / 42.5 | 1.005 [1.003–1.008] | 1.10 | meets |
| `xlsb_eager_cell_value_set_batch1000` | 54,198,536 → 53,915,707 (0.995) | 0.979 | 1,133.4 / 1,791.0 | 1.584 [1.573–1.593] | 2.0 | meets |
| `xlsb_eager_cell_remove_batch1000` | 36,274,413 → 36,678,426 (1.011) | 1.013 | 753.6 / 1,285.4 | 1.707 [1.701–1.723] | 2.0 | meets |
| `xlsb_eager_cell_value_set_single` | 19,537,707 → 18,818,306 (0.963) | 0.974 | 950.1 / 954.1 | 1.002 [0.996–1.011] | 1.10 | meets |
| `xlsb_default_open` | 1,553,636 → 1,607,352 (1.035) | 1.000 | 66.9 / 50.1 | 0.748 [0.741–0.755] | 1.05 or +20 µs | meets |
| `xlsb_eager_cell_remove_single` | 20,636,032 → 20,081,331 (0.973) | 0.974 | 984.8 / 987.8 | 1.001 [0.996–1.009] | 1.10 | meets |
| `xlsb_eager_open` | 2,687,326 → 2,474,076 (0.921) | 0.998 | 65.7 / 77.0 | 1.173 [1.167–1.182] | 1.05 or +20 µs | meets |
| `xlsb_eager_metadata_title_set` | 847,714 → 871,769 (1.028) | 1.043 | 33.4 / 33.8 | 1.016 [1.008–1.020] | 1.10 | meets |
| `xlsb_eager_save_first` | 1,419,093 → 1,400,008 (0.987) | 1.008 | 28.6 / 28.7 | 1.009 [0.994–1.017] | 1.10 | meets |
| `xlsb_eager_save_second` | 1,419,012 → 1,399,154 (0.986) | 1.007 | 23.4 / 23.5 | 1.003 [0.995–1.007] | 1.10 | meets |

### W5 DOC/XLS/RTF

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `doc_default_open` | 18,290,033 → 18,293,554 (1.000) | 1.001 | 849.6 / 854.7 | 1.009 [0.996–1.017] | 1.05 or +20 µs | meets |
| `doc_eager_paragraph_text_set_single` | 38,465,781 → 38,334,139 (0.997) | 0.994 | 1,434.9 / 1,450.1 | 1.004 [0.993–1.017] | 1.10 | meets |
| `doc_eager_open` | 18,295,798 → 18,298,483 (1.000) | 1.000 | 847.5 / 852.1 | 1.010 [0.994–1.018] | 1.05 or +20 µs | meets |
| `doc_eager_metadata_title_set` | 21,947,188 → 21,321,062 (0.971) | 0.972 | 899.1 / 901.8 | 1.003 [0.998–1.010] | 1.10 | meets |
| `doc_eager_save_first` | 438,959 → 261,772 (0.596) | 1.000 | 2.4 / 2.4 | 1.008 [1.000–1.019] | 1.10 | meets |
| `doc_eager_save_second` | 437,642 → 262,517 (0.600) | 1.004 | 1.2 / 1.2 | 0.994 [0.979–1.008] | 1.10 | meets |
| `doc_eager_paragraph_text_get` | 3,793 → 3,779 (0.996) | 1.000 | 25.1 / 0.2 | 0.008 [0.008–0.009] | 1.10 (read) | meets |
| `xls_eager_cell_value_set_batch1000` | 175,874,442 → 179,280,051 (1.019) | 1.002 | 3,948.0 / 4,042.6 | 1.024 [1.018–1.039] | 2.0 | meets |
| `xls_eager_cell_remove_batch1000` | 135,834,812 → 137,564,847 (1.013) | 0.999 | 3,166.9 / 3,288.7 | 1.039 [1.037–1.043] | 2.0 | meets |
| `xls_file_cell_value_set_batch1000` | 168,887,330 → 172,016,153 (1.019) | 1.001 | 3,742.1 / 4,348.4 | 1.170 [1.146–1.186] | 2.0 | meets |
| `xls_file_cell_remove_batch1000` | 133,892,157 → 135,742,551 (1.014) | 0.998 | 3,168.9 / 3,751.0 | 1.187 [1.183–1.203] | 2.0 | meets |
| `xls_eager_cell_value_set_single` | 23,587,691 → 23,481,925 (0.996) | 0.999 | 708.9 / 711.1 | 1.003 [1.000–1.011] | 1.10 | meets |
| `xls_eager_cell_remove_single` | 22,848,698 → 22,809,854 (0.998) | 1.003 | 883.8 / 888.8 | 1.004 [1.002–1.008] | 1.10 | meets |
| `xls_file_cell_value_set_single` | 18,036,911 → 17,805,567 (0.987) | 0.987 | 455.5 / 459.8 | 1.015 [1.005–1.019] | 1.10 | meets |
| `xls_file_cell_remove_single` | 17,397,059 → 17,023,038 (0.979) | 0.992 | 615.8 / 622.8 | 1.013 [1.007–1.017] | 1.10 | meets |
| `xls_default_open` | 5,799,830 → 5,900,050 (1.017) | 1.017 | 236.2 / 238.4 | 1.009 [0.973–1.025] | 1.05 or +20 µs | meets |
| `xls_eager_open` | 5,808,160 → 5,901,562 (1.016) | 1.016 | 235.5 / 234.4 | 0.999 [0.978–1.023] | 1.05 or +20 µs | meets |
| `xls_eager_metadata_title_set` | 7,339,446 → 7,255,523 (0.989) | 0.988 | 263.3 / 263.8 | 1.003 [0.992–1.022] | 1.10 | meets |
| `xls_eager_save_first` | 64,929 → 65,278 (1.005) | 1.003 | 1.2 / 1.3 | 1.041 [1.004–1.061] | 1.10 | meets |
| `xls_eager_save_second` | 60,153 → 60,267 (1.002) | 1.001 | 0.5 / 0.6 | 1.032 [0.968–1.147] | 1.10 | meets |
| `rtf_default_paragraph_text_set_single` | 7,835,657 → 7,817,157 (0.998) | 0.998 | 254.1 / 253.8 | 0.999 [0.988–1.014] | 1.10 | meets |
| `rtf_default_paragraph_text_set_batch1000` | 273,795,522 → 273,526,166 (0.999) | 1.000 | 4,812.7 / 5,185.0 | 1.081 [1.068–1.089] | 2.0 | meets |
| `rtf_default_open` | 4,807,138 → 4,816,140 (1.002) | 1.001 | 170.6 / 172.3 | 1.009 [0.993–1.025] | 1.05 or +20 µs | meets |
| `doc_eager_paragraph_text_set_batch1000` | 53,002,270,241 → 55,192,516,744 (1.041) | 1.041 | — | 1.000 (Ir) | 2.0 | meets |

### W6 ODF

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `odt_default_paragraph_text_set_batch1000` | 59,259,352 → 63,219,198 (1.067) | 1.088 | 1,120.5 / 1,775.7 | 1.580 [1.561–1.612] | 2.0 | meets, >5 % Ir |
| `odt_default_paragraph_text_set_single` | 16,455,584 → 19,024,508 (1.156) | 1.155 | 496.8 / 496.9 | 0.998 [0.993–1.006] | 1.10 | meets, >5 % Ir |
| `odt_default_open` | 2,782,585 → 5,252,697 (1.888) | 1.890 | 149.1 / 149.6 | 1.004 [1.000–1.006] | 1.05 or +20 µs | meets, >5 % Ir |
| `oth_default_paragraph_text_set_batch1000` | 54,080,870 → 56,294,320 (1.041) | 1.054 | 997.2 / 1,638.9 | 1.644 [1.615–1.678] | 2.0 | meets, >5 % Ir |
| `oth_default_paragraph_text_set_single` | 12,945,538 → 14,353,085 (1.109) | 1.108 | 381.2 / 381.7 | 0.999 [0.992–1.011] | 1.10 | meets, >5 % Ir |
| `odm_default_paragraph_text_set_batch1000` | 54,000,204 → 56,018,383 (1.037) | 1.048 | 994.6 / 1,641.7 | 1.649 [1.615–1.682] | 2.0 | meets |
| `odm_default_paragraph_text_set_single` | 12,921,197 → 14,178,614 (1.097) | 1.097 | 380.1 / 379.5 | 0.999 [0.993–1.009] | 1.10 | meets, >5 % Ir |
| `ods_default_cell_value_set_batch1000` | 197,292,202 → 203,062,112 (1.029) | 1.032 | 4,321.0 / 5,107.5 | 1.180 [1.177–1.188] | 2.0 | meets |
| `ods_default_cell_value_set_single` | 186,316,998 → 194,061,526 (1.042) | 1.043 | 6,813.1 / 6,811.6 | 1.000 [0.998–1.001] | 1.10 | meets |
| `ods_default_open` | 34,270,797 → 35,901,827 (1.048) | 1.049 | 1,341.1 / 1,340.5 | 1.000 [0.996–1.004] | 1.05 or +20 µs | meets |
| `odp_default_paragraph_text_set_batch1000` | 110,590,406 → 90,647,522 (0.820) | 0.750 | 1,362.9 / 2,235.6 | 1.641 [1.628–1.661] | 2.0 | meets |
| `odp_default_paragraph_text_set_single` | 46,610,249 → 33,285,489 (0.714) | 0.709 | 1,058.0 / 1,059.4 | 1.001 [0.993–1.008] | 1.10 | meets |
| `odp_default_open` | 2,390,515 → 2,457,616 (1.028) | 1.023 | 62.3 / 64.1 | 1.029 [1.025–1.033] | 1.05 or +20 µs | meets |
| `odg_default_shape_text_set_single` | 415,055,486 → 415,174,360 (1.000) | 0.999 | 13,361.7 / 13,373.2 | 1.001 [0.997–1.003] | 1.10 | meets |
| `odg_default_open` | 229,222,687 → 230,541,204 (1.006) | 1.000 | 6,992.0 / 7,230.7 | 1.034 [1.030–1.036] | 1.05 or +20 µs | meets |
| `odc_default_series_label_set_batch1000` | 270,133,458 → 273,713,092 (1.013) | 1.015 | 5,522.7 / 5,914.1 | 1.069 [1.062–1.077] | 2.0 | meets |
| `odc_default_series_label_set_single` | 129,653,342 → 131,771,953 (1.016) | 1.017 | 4,228.1 / 4,271.8 | 1.009 [1.003–1.015] | 1.10 | meets |
| `odc_default_open` | 38,104,624 → 39,213,750 (1.029) | 1.028 | 1,254.8 / 1,324.9 | 1.057 [1.053–1.064] | 1.05 or +20 µs | miss |
| `odi_default_frame_property_set_batch6` | 10,170,266 → 10,302,295 (1.013) | 1.014 | 171.2 / 175.6 | 1.026 [1.023–1.034] | scaling | meets |
| `odb_default_query_command_set_single` | 74,416,361 → 74,629,669 (1.003) | 1.004 | 2,295.3 / 2,295.5 | 1.000 [0.997–1.003] | 1.10 | meets |
| `odb_default_query_command_set_batch1000` | 40,318,006,497 → 40,477,654,136 (1.004) | 1.001 | — | 1.003 (Ir) | 2.0 | meets |
| `odf_default_token_text_set_single` | 45,227,065 → 45,395,784 (1.004) | 1.003 | 1,583.9 / 1,575.7 | 0.993 [0.983–1.005] | 1.10 | meets |
| `odf_default_token_text_set_batch1000` | 33,829,589,532 → 33,980,952,071 (1.004) | 1.004 | — | 1.000 (Ir) | 2.0 | meets |
| `odg_default_shape_text_set_batch1000` | 168,231,039,273 → 168,259,599,708 (1.000) | 0.997 | — | 1.001 (Ir) | 2.0 | meets |

### W7 iWork

| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |
|---|---|---:|---|---|---|---|
| `pages_default_paragraph_text_set_batch1000` | 27,979,262 → 27,085,494 (0.968) | 0.948 | 388.8 / 582.7 | 1.505 [1.455–1.528] | 2.0 | meets |
| `pages_default_paragraph_text_set_single` | 12,039,561 → 11,228,879 (0.933) | 0.930 | 284.7 / 285.3 | 1.002 [0.990–1.009] | 1.10 | meets |
| `pages_default_open` | 2,391,718 → 2,380,402 (0.995) | 0.982 | 70.4 / 75.3 | 1.069 [1.064–1.073] | 1.05 or +20 µs | meets |
| `keynote_default_paragraph_text_set_batch1000` | 97,352,944 → 93,587,795 (0.961) | 0.951 | 4,758.1 / 5,178.3 | 1.089 [1.084–1.095] | 2.0 | meets |
| `keynote_default_paragraph_text_set_single` | 52,968,614 → 50,399,610 (0.951) | 0.952 | 4,717.7 / 4,739.0 | 1.004 [1.002–1.007] | 1.10 | meets |
| `keynote_default_open` | 32,318,662 → 30,779,901 (0.952) | 0.943 | 1,466.7 / 1,514.4 | 1.033 [1.030–1.036] | 1.05 or +20 µs | meets |
| `numbers_default_cell_value_set_batch1000` | 33,589,587 → 34,044,638 (1.014) | 1.000 | 732.9 / 987.9 | 1.344 [1.316–1.362] | 2.0 | meets |
| `numbers_default_cell_value_set_single` | 3,354,260 → 2,803,867 (0.836) | 0.836 | 153.8 / 156.4 | 1.016 [1.008–1.022] | 1.10 | meets |
| `numbers_default_open` | 1,428,283 → 1,386,556 (0.971) | 0.937 | 56.7 / 64.1 | 1.131 [1.124–1.135] | 1.05 or +20 µs | meets |

### 4.8 W8 File and cross-cutting

- **1k-op batch budget (2.0)**: met by all 21 formats (batch rows above):
  worst unified/leaf XLSX 1.76, ODP 1.67, ODM 1.65, OTH 1.64, XLSB removal
  1.72; DOC, ODG, ODB and ODF formula 1.000 in instructions.
- **Second save**: the second `to_bytes` of one edited package costs no more
  than the baseline's first save on every format with a save route (DOCX
  36.4 / 38.0 µs, XLSX 42.5 / 46.8, PPTX 69.2 / 86.4, XLSB 23.5 / 30.5, DOC
  1.2 / 4.3, XLS 0.5 / 1.3, PPT 1.9 / 7.6; candidate second save / baseline
  first save, unified leg, wall).
- **Python** (`python_bench.py`, release `litchi_py`, A-B-B-A processes;
  no budget): candidate/baseline XLSX `open_bytes` 1.038, `set` one cell
  0.946, `batch` of 1k cell sets 1.083 (the XLSX snapshot bound of §5.3 on
  the default source-backed owner), `get` 1.001; DOCX `open_bytes` 1.010,
  `set` 1.012, `to_bytes` 1.093 (0.7 → 0.8 µs).
- **CLI**: not re-measured (§9).

### 4.9 Existing performance-program scenarios (perf-baseline default matrix)

41 cases, 213 rows, both builds, eight A-B-B-A processes. Instructions per
sample (`perf-baseline-callgrind.tsv`): 39 cases within 1.042;
`xlsx_source_first_cell` 1.051 and `xlsx_source_narrow_column_range_scan`
1.045 count untimed per-sample setup — their timed wall intervals are
0.93–0.97 of the baseline; `odp_existing_append_lifecycle` 1.041 (deferred
by owner, wall 1.056–1.058). Wall rows above 5 % at equal instructions are
the codegen residual of §5.5.

### 4.10 W6 probes

`odf_w6_probe` (the record 0358 routes, unified only, before/after):

| Route | Instructions per round, base → cand | Ratio | Wall ratio | Verdict |
|---|---|---:|---:|---|
| ODT open + get | 18.19 M → 27.05 M | 1.487 | 1.496 | deferred by owner |
| ODT set | 61.85 M → 71.49 M | 1.156 | 1.183 | deferred by owner |
| ODT batch, 1k sets | 36.76 M → 38.86 M | 1.057 | 1.054 | deferred by owner |
| flat ODT open + get | 46.71 M → 71.46 M | 1.530 | 1.225 | deferred by owner |
| ODS open + get | 129.77 M → 135.67 M | 1.045 | 1.069 | meets (Ir) |
| ODS set | 2.98 M → 2.37 M | 0.796 | 0.691 | meets |
| ODS batch, 1k sets | 98.26 M → 101.15 M | 1.029 | 1.048 | meets |
| ODP open + get | 15.39 M → 15.41 M | 1.001 | 1.007 | meets |
| ODP set | 56.44 M → 44.02 M | 0.780 | 0.770 | meets |
| ODG open + get | 826.5 M → 828.8 M | 1.003 | 0.999 | meets |
| ODG set | 1,567 M → 1,569 M | 1.001 | 1.002 | meets |
| ODC open + get | 101.14 M → 109.14 M | 1.079 | 1.079 | explained residual (§5.4) |
| ODC set | 117.93 M → 123.77 M | 1.050 | 1.052 | meets (1.0496) |
| flat ODS open + get | 224.19 M → 229.37 M | 1.023 | 1.015 | meets |

The earlier ODT set 1.070 and flat ODT open 1.090 of record 0358a are now
1.18 and 1.23 against the pre-program baseline (the full W6 + raw-XML-limit
program rather than one wave step); the instruction counts show added work
(§5.1), not codegen. `odf_source_backed_probe` (record 0358a's routes; the
baseline twin `results/0358a/probe_before.rs` built on `839212fa9d` with the
then-current ODP type name): ODS 1 cell 1.058, 100 cells 1.037, DDE 1.004,
sheet metadata 1.012 (wall); **ODP tail append 0.983, peak heap 0.58 → 0.58
MiB (1.000)**; every output SipHash identical across builds.

## 5. Causes of what remains

### 5.1 ODT engine (deferred by owner)

`odt_default_open` 2.78 M → 5.25 M Ir: the baseline's `OpenParse::run`
drove a plain `quick_xml::Reader` with the gated `BindingTracker`
(changes 0221/0224) and resolved namespaces only at shallow depth; the
raw-XML-limit work (0362ul44) replaced it with `ResolvedReader`, which
validates and resolves every event (`resolve_event` and `resolve_prefix`,
+7 M of the 8.4 M per call of the probe). ODT set, OTH and ODM sets share
the parse. Not changed here (owner decision).

### 5.2 OOXML metadata commit

`litchi_ooxml_common::properties::transaction` (record 0362o) reads the
strict graph of the three property families (`Snapshot::read`,
`family_state`), plans an exact reversible patch with an authenticated
scope (`Plan::finish`, `Patch::check`) and applies it to a package clone —
+63 k Ir per DOCX title set (97 k → 162 k), +31 k XLSX, +115 k PPTX: a fixed
1–3 µs per commit, not proportional to the document; every wall
unified/leaf ratio meets 1.10. Follow-up **0377-F2**: read only the changed
family's payload and keep the others' scope as byte references.

### 5.3 XLSX unified cell edits (contract snapshot bound)

The unified XLSX edit is bounded (`EditOptions::with_max_retained_bytes`,
contract §2.5); the baseline had no bound and the 0337 leaf leg has none.
After §3 the bound costs one ledger admission per staged cell
(`preflight_cell_capture_with_facet` 4.1 M + ledger `BTreeMap` 3.7 M per
1,000 cells) and, for a single set, three `workbook_retained_bytes` walks
(≈11 k Ir). The leaf edit itself is +6.2 % (typed-cell and capture plumbing
of W4). Budgets: 1k batch 1.76 (2.0 met); single set 1.081 Ir, wall 1.108
default / 1.095 with `codegen-units=1`. Follow-up **0377-F3**: a flat
per-sheet ledger, and one endpoint walk per commit.

### 5.4 ODC open (odfdom fixture)

After the namespace fast path: duplicate-attribute checks (+7.3 M per 4
rounds), `ResolvedReader` push/resolve (+13 M), UTF-8 validation (+5.7 M) —
the raw-XML admission of 0362ul44 over the fixture's many namespace
declarations. The unified ODC open also builds the object tree at open (the
object formats' open-time structural validation; 2.9 M of 39 M Ir per open on
the builder fixture), which was the pre-existing unified/leaf 1.071
(baseline) / 1.072 (candidate), +70 µs. Follow-up **0377-F4**: fuse the
duplicate check into the resolver's attribute pass; **0377-F5**: build the
tree from the leaf's parsed chart instead of a second walk.

### 5.5 Wall clock above instructions: code generation

perf-baseline `opc_source_open`/`opc_open_owned` (many-small corpora)
+7.7–9.1 % and `xls_fresh_write_to` (xls-large) +9 % wall, in three
separate A-B-B-A sets, with instructions +2.1 % / +2.5 % / +0.1 %.
Cachegrind of `opc_source_open`: I1 misses ×2.04 per sample, conditional
mispredicts ×1.12, data misses level — hot small functions
(`PackURI::new`, `hash_one`, `inspect_element`) moved apart. Rebuilt with
`codegen-units=1` (both trees, same sources), the same cases measure 0.995 –
1.011 (`perf-baseline-rerun.tsv`), as record 0358a found for the ODF
probes. The harness rows with the same pattern (ODS open 1.062 → 1.022,
PPTX metadata 1.071 → 1.007, DOCX save 1.057 → 1.045 with one codegen unit;
`abba-codegen-units-1.json`) are this residual. No code change can target
it; the instruction counts are the regression evidence.

## 6. Bench matrix

`tools/unified_bench_matrix.py` gains `p0377`: per route an after/before row
per leg (instructions, budget 1.05), an instruction unified/leaf row (no
budget) and a wall unified/leaf row (the route's budget); the `odf_w6_probe`
and perf-baseline after/before rows (instructions); the ODP peak-heap rows;
the Python after/before rows (no budget). Remaining misses carry the
explanations of §5 (follow-ups 0377-F2–F5, codegen residual, or "deferred by
owner").

## 7. Gates

Run on the merged tree with the rules' build environment
(`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`):

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-odf-common -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features doc / ppt / xlsx / pptx / odc] --all-targets -- -D warnings` | 0 (all six) |
| `cargo test --all-features -p litchi-pptx` / `-p litchi-xlsx` / `-p litchi-xlsb` / `-p litchi-odf-common` / `-p litchi-odc` | 0 each |
| `cargo test -p litchi-ops --all-features`: `--lib` (1,022), `--doc` (217), all 167 `--test` targets in six chunks (capability and registry goldens unchanged, no re-bless) | 0 |
| `cargo test -p litchi --all-features` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pptx -p litchi-xlsx -p litchi-xlsb -p litchi-odf-common -p litchi-ops` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (71 profiles) |
| tool unit tests (the five required modules plus `test_unified_bench_matrix`; 1,288 tests) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` | 0 each |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `python3 tools/check_crate_boundaries.py` | **1** — two findings in `crates/litchi-keynote/src/package/slide_drawable/named_styles.rs` (392, 548: a private `fn rename` read as `fs::rename`), from trunk commit `af21c5ea69` (iwa-w1c, merged here); not touched by this record |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets`; same for `unified-fin-perf` | 0, 0 |
| perf-baseline default matrix (41 cases, 213 rows) on the candidate | runs again (was exit 1 on trunk, §3) |

## 8. Evidence

Committed (`results/0377/`, 160 KB): `callgrind.tsv`, `abba.json`,
`abba-codegen-units-1.json`, `odf-w6-callgrind.tsv`, `odf-w6-wall.json`,
`odf-sb-wall.json`, `perf-baseline-callgrind.tsv`, `perf-baseline-wall.tsv`,
`perf-baseline-rerun.tsv`, `python.json`, `binaries.sha256`, and the scripts
`build.sh`, `abba.sh`, `callgrind-*.sh`, `summarize.py`, `table.py`,
`python_bench.py`. Raw callgrind/cachegrind outputs and per-run JSON stay in
`/home/zhuhe/code/litchi-worktrees/.coord/fin-perf/` (regenerate with the
scripts against the binaries of `binaries.sha256`).

## 9. Gaps and follow-ups

- **0377-F1** — DOC set: instructions put unified/leaf at 1.18 (baseline
  1.178) — one extra `Document` parse in the unified deferred body open,
  cause not isolated — while wall measures 0.999; not a wall miss.
- **0377-F2** OOXML metadata commit (§5.2); **0377-F3** XLSX ledger (§5.3);
  **0377-F4/F5** ODC open (§5.4).
- ODT/ODP: deferred by owner (§5.1, §4.10).
- CLI rows were not re-measured (no CLI route changed in §3).
- Wall numbers were taken at host load 15–30; the instruction counts are
  unaffected.
