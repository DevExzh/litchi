# DOCX SVG lifecycle dominant phase observations

These are p50 phase maxima within each named lane. They identify where
the scaffold spends time; they are not a causal attribution or speedup
claim. Fixture setup and caller-owned payload construction are outside
the timed operation.

mode=full

| lane | dominant phase | p50 ns | capture | stage | commit | publish | reopen | inverse reopen | inverse | payload | validation | readback |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| native_svg_capture | capture_ns | 160331 | 160331 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 20 | 0 |
| native_floating_capture | capture_ns | 241371 | 241371 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 20 | 0 |
| lazy_inventory_1 | validation_ns | 35540 | 30380 | 0 | 0 | 0 | 0 | 0 | 0 | 1250 | 35540 | 0 |
| lazy_inventory_64 | capture_ns | 914853 | 914853 | 0 | 0 | 0 | 0 | 0 | 0 | 1990 | 388682 | 0 |
| single_attach_1 | publish_ns | 208591 | 35280 | 48000 | 180541 | 208591 | 82920 | 0 | 0 | 0 | 52821 | 0 |
| single_attach_16 | commit_ns | 1066394 | 251981 | 361092 | 1066394 | 981514 | 402911 | 0 | 0 | 0 | 148531 | 0 |
| single_attach_64 | commit_ns | 3797064 | 936244 | 1347975 | 3797064 | 3345383 | 1357985 | 0 | 0 | 0 | 423702 | 0 |
| single_detach_1 | publish_ns | 214951 | 45201 | 43270 | 180231 | 214951 | 65820 | 0 | 0 | 0 | 43400 | 0 |
| single_detach_16 | commit_ns | 1735346 | 354861 | 490352 | 1735346 | 1620316 | 585392 | 0 | 0 | 0 | 250631 | 0 |
| single_detach_64 | commit_ns | 6601935 | 1331415 | 1905058 | 6601935 | 5968243 | 2147869 | 0 | 0 | 0 | 854143 | 0 |
| batch_attach_1 | publish_ns | 205961 | 35160 | 47841 | 178151 | 205961 | 83230 | 0 | 0 | 0 | 51980 | 0 |
| batch_attach_16 | stage_ns | 7509379 | 294611 | 7509379 | 1648076 | 1548946 | 651283 | 0 | 0 | 0 | 263631 | 0 |
| batch_attach_64 | stage_ns | 116064677 | 1317055 | 116064677 | 3355423 | 0 | 0 | 0 | 0 | 0 | 411522 | 3601514 |
| batch_detach_1 | publish_ns | 212340 | 45350 | 43661 | 179190 | 212340 | 65781 | 0 | 0 | 0 | 43110 | 0 |
| batch_detach_16 | stage_ns | 6880197 | 436501 | 6880197 | 1548036 | 1531126 | 403322 | 0 | 0 | 0 | 142221 | 0 |
| batch_detach_64 | stage_ns | 106387191 | 1350238 | 106387191 | 4920148 | 4476846 | 1069787 | 0 | 0 | 0 | 64740 | 0 |
| shared_svg_cleanup | stage_ns | 106211829 | 1350115 | 106211829 | 4909218 | 4466407 | 1065224 | 0 | 0 | 0 | 109091 | 0 |
| exact_inverse_single_1 | publish_ns | 206131 | 34210 | 47940 | 178321 | 206131 | 81960 | 58320 | 44540 | 0 | 89570 | 0 |
| exact_inverse_batch_64 | stage_ns | 115573743 | 1313605 | 115573743 | 3359854 | 0 | 0 | 0 | 0 | 0 | 413441 | 3606564 |
| large_unchanged_media_managed_cap | validation_ns | 1980388 | 33710 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 1980388 | 0 |
| noop_detach_64 | reopen_ns | 1299806 | 922703 | 420 | 1730 | 942013 | 1299806 | 0 | 0 | 0 | 396482 | 0 |
