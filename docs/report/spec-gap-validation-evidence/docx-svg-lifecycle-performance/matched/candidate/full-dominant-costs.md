# DOCX SVG lifecycle dominant phase observations

These are p50 phase maxima within each named lane. They identify where
the scaffold spends time; they are not a causal attribution or speedup
claim. Fixture setup and caller-owned payload construction are outside
the timed operation.

mode=full

| lane | dominant phase | p50 ns | capture | stage | commit | publish | reopen | inverse reopen | inverse | payload | validation | readback |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| native_svg_capture | capture_ns | 168711 | 168711 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 20 | 0 |
| native_floating_capture | capture_ns | 243451 | 243451 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 20 | 0 |
| lazy_inventory_1 | validation_ns | 36070 | 31000 | 0 | 0 | 0 | 0 | 0 | 0 | 1270 | 36070 | 0 |
| lazy_inventory_64 | capture_ns | 916954 | 916954 | 0 | 0 | 0 | 0 | 0 | 0 | 1960 | 394351 | 0 |
| single_attach_1 | publish_ns | 212221 | 35940 | 3210 | 160441 | 212221 | 84461 | 0 | 0 | 0 | 52860 | 0 |
| single_attach_16 | publish_ns | 984875 | 253321 | 4040 | 885124 | 984875 | 404742 | 0 | 0 | 0 | 149841 | 0 |
| single_attach_64 | publish_ns | 3313074 | 935024 | 4970 | 3117543 | 3313074 | 1353786 | 0 | 0 | 0 | 424362 | 0 |
| single_detach_1 | publish_ns | 216061 | 45961 | 930 | 164840 | 216061 | 66421 | 0 | 0 | 0 | 43470 | 0 |
| single_detach_16 | publish_ns | 1618457 | 354572 | 1310 | 1487386 | 1618457 | 584382 | 0 | 0 | 0 | 252401 | 0 |
| single_detach_64 | publish_ns | 5908375 | 1321986 | 1880 | 5592484 | 5908375 | 2135999 | 0 | 0 | 0 | 856874 | 0 |
| batch_attach_1 | publish_ns | 207021 | 35610 | 3170 | 156571 | 207021 | 83521 | 0 | 0 | 0 | 52210 | 0 |
| batch_attach_16 | publish_ns | 1553407 | 295242 | 188780 | 1369006 | 1553407 | 652553 | 0 | 0 | 0 | 263801 | 0 |
| batch_attach_64 | readback_ns | 3619965 | 1317826 | 2720952 | 2278250 | 0 | 0 | 0 | 0 | 0 | 414181 | 3619965 |
| batch_detach_1 | publish_ns | 215581 | 46140 | 1050 | 163071 | 215581 | 66800 | 0 | 0 | 0 | 43650 | 0 |
| batch_detach_16 | publish_ns | 1521037 | 433212 | 14940 | 1347286 | 1521037 | 402281 | 0 | 0 | 0 | 143691 | 0 |
| batch_detach_64 | publish_ns | 4420559 | 1339555 | 74790 | 4146278 | 4420559 | 1049804 | 0 | 0 | 0 | 62320 | 0 |
| shared_svg_cleanup | publish_ns | 4410269 | 1343126 | 74970 | 4134607 | 4410269 | 1049235 | 0 | 0 | 0 | 106581 | 0 |
| exact_inverse_single_1 | publish_ns | 210821 | 35150 | 3390 | 158931 | 210821 | 84440 | 59741 | 46160 | 0 | 90741 | 0 |
| exact_inverse_batch_64 | readback_ns | 3611106 | 1317846 | 2708482 | 2264830 | 0 | 0 | 0 | 0 | 0 | 411002 | 3611106 |
| large_unchanged_media_managed_cap | validation_ns | 1951668 | 33120 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 1951668 | 0 |
| noop_detach_64 | reopen_ns | 1312115 | 928025 | 480 | 3440 | 943574 | 1312115 | 0 | 0 | 0 | 402772 | 0 |
