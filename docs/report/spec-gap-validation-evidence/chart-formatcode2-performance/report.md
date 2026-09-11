# `formatcode2` bounded allocator/runtime profile

This report contains absolute current-source measurements from three fresh processes and twenty measured samples per lane across 28 bounded lanes. The timer covers the named `formatcode2` owner operation with fixture construction and post-timer semantic checks outside the timed interval. `write_to` timing includes the sink's nonallocating byte-count, FNV-1a checksum, and write-call counting work; it makes no comparison claim against the Vec-returning lane. Requested allocation bytes and incremental peak live bytes come from a process-local counting allocator; RSS is whole-process `/usr/bin/time -v` RSS.

| lane | fresh processes | samples | input bytes | p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS KiB |
|---|---:|---:|---:|---:|---:|---:|---:|
| element_small_read | 3 | 60 | 180 | 2710 / 2900 / 3000 | 2142 / 2142 | 755 / 755 | 13808–13856 |
| element_small_noop | 3 | 60 | 180 | 70 / 90 / 100 | 180 / 180 | 180 / 180 | 13780–13856 |
| element_small_scalar_edit | 3 | 60 | 180 | 300 / 360 / 410 | 397 / 397 | 397 / 397 | 13796–13860 |
| element_small_clone | 3 | 60 | 180 | 50 / 50 / 60 | 0 / 0 | 0 / 0 | 13776–13852 |
| element_small_read_shared | 3 | 60 | 180 | 2720 / 3080 / 6340 | 1942 / 1942 | 755 / 755 | 13788–13856 |
| element_small_write_to | 3 | 60 | 180 | 210 / 210 / 220 | 0 / 0 | 0 / 0 | 13776–13860 |
| element_small_malformed | 3 | 60 | 106 | 2100 / 2330 / 2440 | 1580 / 1580 | 755 / 755 | 13776–13920 |
| element_near_limit_read | 3 | 60 | 1048575 | 963444 / 979764 / 984173 | 4457744 / 4457744 | 2031897 / 2031897 | 13792–13952 |
| element_near_limit_noop | 3 | 60 | 1048575 | 110200 / 118170 / 120140 | 1048575 / 1048575 | 1048575 / 1048575 | 14760–14884 |
| element_near_limit_scalar_edit | 3 | 60 | 1048575 | 346802 / 361771 / 371041 | 1179759 / 1179759 | 1179759 / 1179759 | 14776–14880 |
| element_near_limit_clone | 3 | 60 | 1048575 | 50 / 50 / 80 | 0 / 0 | 0 / 0 | 13852–13860 |
| element_near_limit_read_shared | 3 | 60 | 1048575 | 945024 / 957314 / 971334 | 3409152 / 3409152 | 2031897 / 2031897 | 13784–13976 |
| element_near_limit_write_to | 3 | 60 | 1048575 | 936583 / 939673 / 941224 | 0 / 0 | 0 / 0 | 13796–13856 |
| element_near_limit_malformed | 3 | 60 | 1048576 | 876523 / 899544 / 900923 | 3278037 / 3278037 | 2031897 / 2031897 | 13776–13856 |
| attribute_small_read | 3 | 60 | 129 | 2720 / 2910 / 3080 | 2129 / 2129 | 1227 / 1227 | 13804–13856 |
| attribute_small_noop | 3 | 60 | 129 | 80 / 90 / 100 | 129 / 129 | 129 / 129 | 13776–13780 |
| attribute_small_scalar_edit | 3 | 60 | 129 | 160 / 190 / 200 | 148 / 148 | 148 / 148 | 13788–13856 |
| attribute_small_clone | 3 | 60 | 129 | 50 / 60 / 70 | 0 / 0 | 0 / 0 | 13768–13900 |
| attribute_small_read_shared | 3 | 60 | 129 | 2720 / 2940 / 2980 | 1977 / 1977 | 1227 / 1227 | 13792–13856 |
| attribute_small_write_to | 3 | 60 | 129 | 170 / 170 / 170 | 0 / 0 | 0 / 0 | 13856–13856 |
| attribute_small_malformed | 3 | 60 | 102 | 2030 / 2300 / 2350 | 1690 / 1690 | 1018 / 1018 | 13788–13860 |
| attribute_near_limit_read | 3 | 60 | 1048575 | 2076528 / 2087077 / 2092918 | 5400745 / 5400745 | 2228560 / 2228560 | 13776–13852 |
| attribute_near_limit_noop | 3 | 60 | 1048575 | 110951 / 118970 / 129370 | 1048575 / 1048575 | 1048575 / 1048575 | 14820–14888 |
| attribute_near_limit_scalar_edit | 3 | 60 | 1048575 | 360391 / 368631 / 369681 | 1179599 / 1179599 | 1179599 / 1179599 | 14816–14844 |
| attribute_near_limit_clone | 3 | 60 | 1048575 | 50 / 60 / 60 | 0 / 0 | 0 / 0 | 13848–13856 |
| attribute_near_limit_read_shared | 3 | 60 | 1048575 | 2057817 / 2068357 / 2081968 | 4352153 / 4352153 | 2228560 / 2228560 | 13852–13912 |
| attribute_near_limit_write_to | 3 | 60 | 1048575 | 936334 / 940143 / 946903 | 0 / 0 | 0 / 0 | 13812–13860 |
| attribute_near_limit_malformed | 3 | 60 | 1048576 | 2057927 / 2067488 / 2073308 | 6449294 / 6449294 | 2228560 / 2228560 | 13780–13792 |

`element_*` uses a short source-preserving element or a valid source of `MAX_XML_BYTES - 1` bytes. `attribute_*` uses a complete host start tag containing the qualified chart attribute; its near-limit form fills bounded ordinary attributes while keeping every value below the per-attribute ceiling. Each near-limit decoded value is close to `MAX_VALUE_BYTES`; its malformed lane appends one non-whitespace byte and remains within the source ceiling. Each small malformed lane exercises an unpaired ST_Xstring surrogate escape. `clone` is the prepared typed owner's `Clone`, `read_shared` parses an existing `Arc<[u8]>` prepared outside the timer, `no-op` serializes the unchanged prepared owner into an owned `Vec`, `write_to` serializes the same unchanged owner into a counting/hash sink, and `scalar_edit` changes the decoded value and serializes it.

These are scoped absolute observations. They do not establish a before/after speedup, an asymptotic result, a host-placement guarantee, or a whole-library performance claim. The shared attribute owner is measured only for its complete start-tag contract; the package host still owns placement and parent grammar.

Raw per-process JSON, `/usr/bin/time -v` receipts, source/build manifests, exact commands, and source hashes are retained in this directory.
