# DOCX SVG lifecycle dominant phase observations

These are p50 phase maxima within each named lane. They identify where
the scaffold spends time; they are not a causal attribution or speedup
claim. Fixture setup and caller-owned payload construction are outside
the timed operation.

mode=full

| lane | dominant phase | p50 ns | capture | stage | commit | publish | reopen | inverse reopen | inverse | payload | validation | readback |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| native_svg_capture | capture_ns | 162681 | 162681 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 20 | 0 |
| native_floating_capture | capture_ns | 239431 | 239431 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 20 | 0 |
| lazy_inventory_1 | validation_ns | 36020 | 29990 | 0 | 0 | 0 | 0 | 0 | 0 | 1250 | 36020 | 0 |
| lazy_inventory_64 | capture_ns | 896094 | 896094 | 0 | 0 | 0 | 0 | 0 | 0 | 1940 | 393892 | 0 |
| single_attach_1 | publish_ns | 208301 | 35030 | 47150 | 179531 | 208301 | 82930 | 0 | 0 | 0 | 53050 | 0 |
| single_attach_16 | commit_ns | 1055465 | 250181 | 355502 | 1055465 | 972934 | 401841 | 0 | 0 | 0 | 149670 | 0 |
| single_attach_64 | commit_ns | 3771346 | 927663 | 1330456 | 3771346 | 3325094 | 1359086 | 0 | 0 | 0 | 430242 | 0 |
| single_detach_1 | publish_ns | 214991 | 45070 | 42391 | 178461 | 214991 | 66741 | 0 | 0 | 0 | 44220 | 0 |
| single_detach_16 | commit_ns | 1721487 | 351772 | 481922 | 1721487 | 1609607 | 583632 | 0 | 0 | 0 | 253971 | 0 |
| single_detach_64 | commit_ns | 6502667 | 1311786 | 1863908 | 6502667 | 5893955 | 2131239 | 0 | 0 | 0 | 861153 | 0 |
| batch_attach_1 | publish_ns | 206471 | 34610 | 47250 | 177911 | 206471 | 82740 | 0 | 0 | 0 | 52890 | 0 |
| batch_attach_16 | stage_ns | 7409551 | 292981 | 7409551 | 1636186 | 1540677 | 651473 | 0 | 0 | 0 | 266611 | 0 |
| batch_attach_64 | stage_ns | 114497894 | 1316566 | 114497894 | 3335754 | 0 | 0 | 0 | 0 | 0 | 416781 | 3596805 |
| batch_detach_1 | publish_ns | 210611 | 44690 | 42970 | 177261 | 210611 | 65951 | 0 | 0 | 0 | 43750 | 0 |
| batch_detach_16 | stage_ns | 6724498 | 434161 | 6724498 | 1530196 | 1519747 | 400132 | 0 | 0 | 0 | 143701 | 0 |
| batch_detach_64 | stage_ns | 103059045 | 1334486 | 103059045 | 4793790 | 4388248 | 1044914 | 0 | 0 | 0 | 65411 | 0 |
| shared_svg_cleanup | stage_ns | 103515367 | 1326185 | 103515367 | 4804160 | 4367578 | 1047075 | 0 | 0 | 0 | 108650 | 0 |
| exact_inverse_single_1 | publish_ns | 206831 | 34220 | 47510 | 178461 | 206831 | 83240 | 58561 | 44500 | 0 | 90991 | 0 |
| exact_inverse_batch_64 | stage_ns | 115092705 | 1323466 | 115092705 | 3336274 | 0 | 0 | 0 | 0 | 0 | 422422 | 3616975 |
| large_unchanged_media_managed_cap | validation_ns | 1964708 | 32790 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 1964708 | 0 |
| noop_detach_64 | reopen_ns | 1301475 | 914863 | 430 | 1730 | 934344 | 1301475 | 0 | 0 | 0 | 405252 | 0 |
