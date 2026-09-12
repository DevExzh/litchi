# DOCX SVG lifecycle dominant phase observations

These are p50 phase maxima within each named lane. They identify where
the scaffold spends time; they are not a causal attribution or speedup
claim. Fixture setup and caller-owned payload construction are outside
the timed operation.

mode=smoke

| lane | dominant phase | p50 ns | capture | stage | commit | publish | reopen | inverse reopen | inverse | payload | validation | readback |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| native_svg_capture | capture_ns | 383681 | 383681 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 360 | 0 |
| native_floating_capture | capture_ns | 606322 | 606322 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 30 | 0 |
| lazy_inventory_1 | validation_ns | 75850 | 65780 | 0 | 0 | 0 | 0 | 0 | 0 | 2440 | 75850 | 0 |
| lazy_inventory_64 | capture_ns | 976103 | 976103 | 0 | 0 | 0 | 0 | 0 | 0 | 3420 | 433831 | 0 |
| single_attach_1 | publish_ns | 299241 | 76150 | 67991 | 229340 | 299241 | 103721 | 0 | 0 | 0 | 72180 | 0 |
| single_attach_16 | commit_ns | 1144473 | 325301 | 399342 | 1144473 | 1077664 | 426681 | 0 | 0 | 0 | 179621 | 0 |
| single_attach_64 | commit_ns | 3887512 | 975503 | 1369934 | 3887512 | 3424941 | 1372514 | 0 | 0 | 0 | 454952 | 0 |
| single_detach_1 | publish_ns | 274221 | 90711 | 56530 | 217600 | 274221 | 102491 | 0 | 0 | 0 | 61560 | 0 |
| single_detach_16 | commit_ns | 1811536 | 452922 | 490951 | 1811536 | 1724815 | 586402 | 0 | 0 | 0 | 288441 | 0 |
| single_detach_64 | commit_ns | 6755751 | 1369004 | 1943247 | 6755751 | 6117010 | 2190397 | 0 | 0 | 0 | 882372 | 0 |
| batch_attach_1 | publish_ns | 270041 | 71760 | 66430 | 224761 | 270041 | 119760 | 0 | 0 | 0 | 73880 | 0 |
| batch_attach_16 | stage_ns | 7546923 | 329912 | 7546923 | 1661346 | 1670145 | 657622 | 0 | 0 | 0 | 273471 | 0 |
| batch_attach_64 | stage_ns | 116241919 | 1386975 | 116241919 | 3424471 | 0 | 0 | 0 | 0 | 0 | 473531 | 3695162 |
| batch_detach_1 | publish_ns | 274141 | 103580 | 52961 | 235620 | 274141 | 80780 | 0 | 0 | 0 | 60021 | 0 |
| batch_detach_16 | stage_ns | 6941112 | 485391 | 6941112 | 1607795 | 1624765 | 424692 | 0 | 0 | 0 | 158680 | 0 |
| batch_detach_64 | stage_ns | 106813599 | 1388325 | 106813599 | 4949146 | 4626205 | 1082103 | 0 | 0 | 0 | 85481 | 0 |
| shared_svg_cleanup | stage_ns | 105910967 | 1397794 | 105910967 | 4924286 | 4548344 | 1080094 | 0 | 0 | 0 | 118340 | 0 |
| exact_inverse_single_1 | publish_ns | 279131 | 101990 | 73280 | 259851 | 279131 | 102970 | 67050 | 63340 | 0 | 129791 | 0 |
| exact_inverse_batch_64 | stage_ns | 116292360 | 1441134 | 116292360 | 3383061 | 0 | 0 | 0 | 0 | 0 | 456332 | 3683081 |
| large_unchanged_media_managed_cap | validation_ns | 1953427 | 105330 | 0 | 0 | 0 | 0 | 0 | 0 | 0 | 1953427 | 0 |
| noop_detach_64 | reopen_ns | 1311255 | 977063 | 1370 | 3020 | 967923 | 1311255 | 0 | 0 | 0 | 450071 | 0 |
