# XLSX form-control scalar lifecycle performance

Status: **pass**; raw samples: **360**.

Percentiles use linear interpolation over each fixture/lane's retained raw samples.
Latency is reported in nanoseconds; allocation and RSS fields retain their raw units.

| Fixture | Lane | n | Field | elapsed p50 | elapsed p95 | elapsed p99 | peak live bytes p95 | RSS after p95 |
| --- | --- | ---: | --- | ---: | ---: | ---: | ---: | ---: |
| singlecontrol | eager_noop_save_reopen | 15 | Checked | 1820179 | 1828960 | 1831911.2 | 234809 | 9412608 |
| singlecontrol | eager_read | 15 | Checked | 759994 | 770881 | 771827.4 | 66796 | 9412608 |
| singlecontrol | eager_scalar_save_reopen | 15 | Checked | 3985289 | 4000011.7 | 4005297.54 | 8411852 | 9469952 |
| singlecontrol | source_forward_apply | 15 | Checked | 906304 | 917304.7 | 930520.14 | 78601 | 9474048 |
| singlecontrol | source_inverse_apply | 15 | Checked | 872915 | 899990 | 911330 | 80299 | 9474048 |
| singlecontrol | source_noop_save_reopen | 15 | Checked | 2669123 | 2674861 | 2675286.6 | 268672 | 9412608 |
| singlecontrol | source_read | 15 | Checked | 803513 | 810102 | 811423.6 | 140705 | 9412608 |
| singlecontrol | source_scalar_save_reopen | 15 | Checked | 3591387 | 3616298.7 | 3617121.34 | 571957 | 9474048 |
| tdf120301_xmlSpaceParsing | eager_noop_save_reopen | 15 | NoThreeD | 2510712 | 2533066.3 | 2535799.66 | 182378 | 9166848 |
| tdf120301_xmlSpaceParsing | eager_read | 15 | NoThreeD | 1121575 | 1136141.3 | 1150993.06 | 115230 | 9166848 |
| tdf120301_xmlSpaceParsing | eager_scalar_save_reopen | 15 | NoThreeD | 5112055 | 5164444 | 5173980.8 | 8415297 | 9183232 |
| tdf120301_xmlSpaceParsing | source_forward_apply | 15 | NoThreeD | 1226956 | 1244415 | 1247007.8 | 124463 | 9256960 |
| tdf120301_xmlSpaceParsing | source_inverse_apply | 15 | NoThreeD | 1195736 | 1206886 | 1207782 | 126994 | 9256960 |
| tdf120301_xmlSpaceParsing | source_noop_save_reopen | 15 | NoThreeD | 3757488 | 3789792.3 | 3797677.66 | 238764 | 9166848 |
| tdf120301_xmlSpaceParsing | source_read | 15 | NoThreeD | 1198566 | 1214415 | 1217231.8 | 159453 | 9166848 |
| tdf120301_xmlSpaceParsing | source_scalar_save_reopen | 15 | NoThreeD | 4995324 | 5054569.3 | 5096597.86 | 617949 | 9220096 |
| tdf134769 | eager_noop_save_reopen | 15 | NoThreeD | 6875463 | 6916938 | 6923294 | 586372 | 10223616 |
| tdf134769 | eager_read | 15 | NoThreeD | 3269155 | 3281967 | 3288110.2 | 466564 | 10223616 |
| tdf134769 | eager_scalar_save_reopen | 15 | NoThreeD | 17940287 | 18262226.3 | 18272396.46 | 8418219 | 10227712 |
| tdf134769 | source_forward_apply | 15 | NoThreeD | 3371316 | 3413083 | 3415429.4 | 476836 | 10227712 |
| tdf134769 | source_inverse_apply | 15 | NoThreeD | 3357956 | 3393956 | 3405100 | 481663 | 10227712 |
| tdf134769 | source_noop_save_reopen | 15 | NoThreeD | 10294410 | 10343105 | 10349517 | 800984 | 10223616 |
| tdf134769 | source_read | 15 | NoThreeD | 3334987 | 3375054 | 3407399.6 | 468734 | 10223616 |
| tdf134769 | source_scalar_save_reopen | 15 | NoThreeD | 13752957 | 13842509.6 | 13935924.32 | 922891 | 10227712 |
