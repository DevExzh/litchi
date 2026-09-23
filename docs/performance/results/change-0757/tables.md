## latency
| selector | shape | samples | before p50 ms | after p50 ms | window 1 ratio [95% CI] | window 2 ratio [95% CI] | same output |
| --- | --- | ---: | ---: | ---: | ---: | ---: | --- |
| `xls_fresh_write_to` | tiny | 100+1000 | 0.0049 | 0.0050 | **1.015** [1.012, 1.018] | 1.016 [1.000, 1.023] | True |
| `xls_fresh_write_to` | large | 10+100 | 1.0986 | 1.1843 | **1.079** [1.074, 1.083] | 1.053 [0.970, 1.074] | True |
| `xls_fresh_write_to` | payload-heavy | 5+40 | 2.3773 | 2.4078 | **1.013** [1.000, 1.018] | 1.012 [1.010, 1.014] | True |
| `xls_semantic_one_edit_save` | tiny | 40+400 | 0.0496 | 0.0496 | **0.997** [0.995, 1.003] | 0.995 [0.993, 0.998] | True |
| `xls_semantic_one_edit_save` | large | 10+100 | 1.5287 | 1.5164 | **0.991** [0.987, 0.998] | 0.995 [0.985, 0.997] | True |
| `doc_fresh_write_to` | tiny | 100+1000 | 0.0058 | 0.0058 | **0.994** [0.981, 1.003] | 0.987 [0.963, 0.998] | True |
| `doc_fresh_write_to` | large | 20+200 | 0.1425 | 0.1362 | **0.956** [0.954, 0.958] | 0.964 [0.960, 0.972] | True |
| `doc_fresh_write_to` | payload-heavy | 5+40 | 1.6117 | 1.6123 | **1.000** [0.995, 1.008] | 1.001 [0.998, 1.006] | True |
| `ppt_fresh_write_to` | tiny | 100+1000 | 0.0098 | 0.0095 | **0.972** [0.940, 1.025] | 0.942 [0.924, 0.999] | True |
| `ppt_fresh_write_to` | large | 20+200 | 0.0717 | 0.0738 | **1.023** [1.015, 1.087] | 1.024 [1.014, 1.043] | True |
| `ppt_fresh_write_to` | payload-heavy | 5+40 | 0.8060 | 0.6374 | **0.792** [0.775, 0.801] | 0.805 [0.784, 0.810] | True |

## counters
| selector | shape | user instructions | user cycles | kernel instructions | page faults |
| --- | --- | --- | --- | --- | --- |
| `xls_fresh_write_to` | tiny | 82,289 → 84,833 (+3.1%) | 24,296 → 26,001 (+7.0%) | 33,183 → 32,950 (-0.7%) | 0 → 0 |
| `xls_fresh_write_to` | large | 17,649,538 → 17,057,174 (-3.4%) | 4,556,245 → 4,740,057 (+4.0%) | 1,034,411 → 2,347,696 (+127.0%) | 188 → 435 (+131.2%) |
| `xls_fresh_write_to` | payload-heavy | 21,374,377 → 21,390,839 (+0.1%) | 7,648,428 → 7,646,418 (-0.0%) | 11,083,450 → 11,158,970 (+0.7%) | 2,103 → 2,119 (+0.7%) |
| `xls_semantic_one_edit_save` | tiny | 1,534,817 → 1,526,230 (-0.6%) | 517,183 → 503,027 (-2.7%) | 38,639 → 34,572 (-10.5%) | 0 → 0 |
| `xls_semantic_one_edit_save` | large | 163,810,721 → 162,374,080 (-0.9%) | 46,323,702 → 46,460,324 (+0.3%) | 2,889,032 → 2,641,112 (-8.6%) | 516 → 468 (-9.2%) |
| `doc_fresh_write_to` | tiny | 96,337 → 96,036 (-0.3%) | 31,142 → 24,642 (-20.9%) | 33,984 → 33,532 (-1.3%) | 0 → 0 |
| `doc_fresh_write_to` | large | 3,091,954 → 3,005,859 (-2.8%) | 668,747 → 639,218 (-4.4%) | 40,062 → 29,942 (-25.3%) | -0 → 0 |
| `doc_fresh_write_to` | payload-heavy | 8,392,933 → 8,390,049 (-0.0%) | 4,675,016 → 4,642,348 (-0.7%) | 13,291,416 → 13,244,546 (-0.4%) | 2,517 → 2,517 (+0.0%) |
| `ppt_fresh_write_to` | tiny | 178,868 → 178,192 (-0.4%) | 41,850 → 41,597 (-0.6%) | 35,350 → 33,304 (-5.8%) | 0 → 0 |
| `ppt_fresh_write_to` | large | 1,365,533 → 1,373,150 (+0.6%) | 313,807 → 284,270 (-9.4%) | 66,902 → 76,544 (+14.4%) | 2 → 2 (-19.2%) |
| `ppt_fresh_write_to` | payload-heavy | 7,285,594 → 7,085,564 (-2.7%) | 3,704,551 → 3,446,720 (-7.0%) | 638,885 → 598,213 (-6.4%) | 113 → 106 (-5.6%) |

## probe (8 processes, default malloc)

| probe case | writes per process | processes | before p50 ms | after p50 ms | paired ratio [95% CI] | instructions per write (callgrind) | distinct outputs before / after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| `xls_multi_string/distinct` | 30 | 8 | 12.7447 | 14.3159 | **1.124** [1.098, 1.138] | 166,885,036 → 185,916,010 (1.114) | 4 / 1 |
| `xls_multi_string/repeated` | 30 | 8 | 10.5242 | 10.9233 | **1.039** [1.032, 1.061] | 142,834,248 → 162,651,230 (1.139) | 4 / 1 |
| `xls_fresh_write_to/tiny` | 3000 | 8 | 0.0044 | 0.0046 | **1.037** [1.031, 1.045] | 60,519 → 63,851 (1.055) | 1 / 1 |
| `xls_fresh_write_to/large` | 300 | 8 | 0.5618 | 0.5391 | **0.961** [0.954, 0.965] | 9,884,364 → 9,132,825 (0.924) | 1 / 1 |
| `xls_fresh_write_to/payload-heavy` | 40 | 8 | 0.9548 | 0.9524 | **0.997** [0.996, 1.000] | 22,694,088 → 22,699,619 (1.000) | 1 / 1 |
| `doc_fresh_write_to/tiny` | 3000 | 8 | 0.0069 | 0.0068 | **0.991** [0.983, 1.016] | 102,733 → 103,232 (1.005) | 1 / 1 |
| `doc_fresh_write_to/large` | 500 | 8 | 0.0954 | 0.0952 | **0.996** [0.989, 1.001] | 1,742,653 → 1,735,214 (0.996) | 1 / 1 |
| `doc_fresh_write_to/payload-heavy` | 40 | 8 | 0.3359 | 0.3341 | **0.995** [0.994, 0.997] | 6,536,234 → 6,534,306 (1.000) | 1 / 1 |
| `ppt_fresh_write_to/tiny` | 3000 | 8 | 0.0148 | 0.0148 | **1.000** [0.997, 1.001] | 170,597 → 170,804 (1.001) | 1 / 1 |
| `ppt_fresh_write_to/large` | 500 | 8 | 0.0957 | 0.0957 | **1.000** [0.996, 1.002] | 1,034,958 → 1,026,034 (0.991) | 1 / 1 |
| `ppt_fresh_write_to/payload-heavy` | 40 | 8 | 0.4081 | 0.4067 | **0.997** [0.987, 1.000] | 14,042,781 → 14,060,829 (1.001) | 1 / 1 |

## probe multi-string (16 processes, default malloc)

| probe case | writes per process | processes | before p50 ms | after p50 ms | paired ratio [95% CI] | instructions per write (callgrind) | distinct outputs before / after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| `xls_multi_string/distinct` | 30 | 16 | 12.9017 | 14.4698 | **1.122** [1.115, 1.126] | 166,172,723 → 188,615,049 (1.135) | 8 / 1 |
| `xls_multi_string/repeated` | 30 | 16 | 10.4258 | 10.9941 | **1.057** [1.035, 1.081] | 141,101,098 → 160,299,156 (1.136) | 8 / 1 |

## probe multi-string (16 processes, pinned malloc thresholds)

| probe case | writes per process | processes | before p50 ms | after p50 ms | paired ratio [95% CI] | instructions per write (callgrind) | distinct outputs before / after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| `xls_multi_string/distinct` | 30 | 16 | 12.7591 | 12.8149 | **1.004** [1.000, 1.012] | 165,148,406 → 187,023,776 (1.132) | 8 / 1 |
| `xls_multi_string/repeated` | 30 | 16 | 10.2554 | 10.8672 | **1.061** [1.055, 1.064] | 141,085,478 → 162,531,220 (1.152) | 8 / 1 |

## allocations

| probe case | allocations | reallocations | allocated bytes | peak live bytes |
| --- | --- | --- | --- | --- |
| `xls_multi_string/distinct` | 21,829 → 21,834 (+0.0%) | 30,184 → 30,145 (-0.1%) | 29,661,811 → 29,604,408 (-0.2%) | 13,137,049 → 12,772,761 (-2.8%) |
| `xls_multi_string/repeated` | 21,799 → 21,803 (+0.0%) | 30,173 → 30,134 (-0.1%) | 14,447,786 → 14,390,188 (-0.4%) | 5,353,541 → 5,353,541 (+0.0%) |
| `xls_fresh_write_to/tiny` | 61 → 61 (+0.0%) | 30 → 30 (+0.0%) | 15,960 → 15,832 (-0.8%) | 37,268 → 37,268 (+0.0%) |
| `xls_fresh_write_to/large` | 1,214 → 1,214 (+0.0%) | 2,135 → 2,135 (+0.0%) | 1,372,507 → 1,306,966 (-4.8%) | 642,336 → 642,336 (+0.0%) |
| `xls_fresh_write_to/payload-heavy` | 2,029 → 2,030 (+0.0%) | 294 → 294 (+0.0%) | 13,028,977 → 13,023,441 (-0.0%) | 12,909,728 → 12,909,728 (+0.0%) |
| `doc_fresh_write_to/tiny` | 116 → 116 (+0.0%) | 39 → 39 (+0.0%) | 54,585 → 54,585 (+0.0%) | 68,296 → 68,296 (+0.0%) |
| `doc_fresh_write_to/large` | 1,670 → 1,670 (+0.0%) | 93 → 93 (+0.0%) | 458,033 → 458,033 (+0.0%) | 330,486 → 330,486 (+0.0%) |
| `doc_fresh_write_to/payload-heavy` | 576 → 576 (+0.0%) | 74 → 74 (+0.0%) | 15,551,507 → 15,551,507 (+0.0%) | 15,512,216 → 15,512,216 (+0.0%) |
| `ppt_fresh_write_to/tiny` | 241 → 241 (+0.0%) | 152 → 152 (+0.0%) | 46,636 → 46,636 (+0.0%) | 54,384 → 54,384 (+0.0%) |
| `ppt_fresh_write_to/large` | 1,873 → 1,873 (+0.0%) | 731 → 731 (+0.0%) | 354,184 → 354,184 (+0.0%) | 155,350 → 155,350 (+0.0%) |
| `ppt_fresh_write_to/payload-heavy` | 1,937 → 1,937 (+0.0%) | 731 → 731 (+0.0%) | 20,880,212 → 20,880,212 (+0.0%) | 15,580,706 → 15,580,706 (+0.0%) |

## tunables

| selector / shape | malloc thresholds | leg order A B B A: timed p50 ms | page faults per iteration | user instructions per iteration |
| --- | --- | --- | --- | --- |
| `ppt_fresh_write_to-payload-heavy` | default | A 0.687 / B 0.655 / B 0.685 / A 0.668 | 113 / 106 / 106 / 113 | 7.29M / 7.09M / 7.09M / 7.29M |
| `ppt_fresh_write_to-payload-heavy` | pinned | A 0.631 / B 0.686 / B 0.672 / A 0.666 | 6 / 0 / -0 / 6 | 7.10M / 7.35M / 7.35M / 7.10M |
| `xls_fresh_write_to-large` | default | A 1.097 / B 1.180 / B 1.189 / A 1.113 | 208 / 414 / 435 / 230 | 17.69M / 17.05M / 17.08M / 17.65M |
| `xls_fresh_write_to-large` | pinned | A 0.991 / B 0.978 / B 0.976 / A 0.997 | -0 / 0 / 0 / 0 | 17.63M / 17.06M / 17.03M / 17.68M |
