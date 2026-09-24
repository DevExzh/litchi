# XLSB `themeFamily` host profile

This report aggregates 3 fresh processes with 30 measured samples per process for each of 18 lanes.
Elapsed time is measured with the harness's process-local counting allocator installed. It is allocator-instrumented operation time for the stated prepared scope, not a production latency claim.
Requested allocation bytes charge direct allocation sizes plus each successful realloc's new size. Raw JSON records realloc old/new sizes, exact live balance, and peak live as an incremental delta above the closure's live-before baseline; it is not total process or retained source/result memory.
Maximum RSS is whole-process startup/warm-up/sample RSS from `/usr/bin/time`, not per-operation memory. Source checks, semantic hashes, pointer checks, and inverse checks are outside timed/allocation regions.

`metadata_read` is the new eager XLSB host read including family discovery. `source_read` uses a counted `ReadAt` source. `codec_read` is a shared DrawingML Theme codec component reference, while `base_edit` is a host base-Theme edit reference; neither is a HEAD-before baseline or a whole-library comparison.

| fixture | operation | processes | samples | package B | Theme B | family B | allocator-instrumented p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS range | ReadAt calls / requested / returned B | source share |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|---:|---|
| native | codec_read | 3 | 90 | 7566 | 8390 | 261 | 69941 / 75430 / 78290 | 34040 / 34040 | 9728 / 9728 | 4936–5096 KiB | 0 / 0 / 0 | n/a |
| native | metadata_read | 3 | 90 | 7566 | 8390 | 261 | 322511 / 343791 / 353481 | 627155 / 627155 | 130038 / 130038 | 4924–5100 KiB | 0 / 0 / 0 | n/a |
| native | source_read | 3 | 90 | 7566 | 8390 | 261 | 296471 / 316512 / 321291 | 816142 / 816142 | 135443 / 135443 | 4880–5104 KiB | 21 / 4167 / 4167 | n/a |
| native | family_clone | 3 | 90 | 7566 | 8390 | 261 | 40 / 40 / 40 | 0 / 0 | 0 / 0 | 4936–5120 KiB | 0 / 0 / 0 | yes |
| native | noop | 3 | 90 | 7566 | 8390 | 261 | 728643 / 735463 / 740663 | 396762 / 396762 | 30091 / 30091 | 4928–5072 KiB | 0 / 0 / 0 | yes |
| native | add | 3 | 90 | 7566 | 8390 | 261 | 2240828 / 2256919 / 2416589 | 1336802 / 1336802 | 77623 / 77623 | 4944–5168 KiB | 0 / 0 / 0 | n/a |
| native | update | 3 | 90 | 7566 | 8390 | 261 | 2263339 / 2279579 / 2395839 | 1360747 / 1360747 | 79173 / 79173 | 4936–5072 KiB | 0 / 0 / 0 | n/a |
| native | remove | 3 | 90 | 7566 | 8390 | 261 | 2228628 / 2245228 / 2252078 | 1331670 / 1331670 | 77093 / 77093 | 4880–5104 KiB | 0 / 0 / 0 | n/a |
| native | base_edit | 3 | 90 | 7566 | 8390 | 261 | 2029508 / 2044298 / 2053467 | 1191123 / 1191123 | 79175 / 79175 | 4948–5088 KiB | 0 / 0 / 0 | n/a |
| opaque | codec_read | 3 | 90 | 8946 | 35140 | 26949 | 158221 / 184021 / 195941 | 36462 / 36462 | 10190 / 10190 | 4880–4948 KiB | 0 / 0 / 0 | n/a |
| opaque | metadata_read | 3 | 90 | 8946 | 35140 | 26949 | 1041344 / 1061983 / 1079584 | 1036977 / 1036977 | 156162 / 156162 | 4928–4948 KiB | 0 / 0 / 0 | n/a |
| opaque | source_read | 3 | 90 | 8946 | 35140 | 26949 | 979624 / 998714 / 1007364 | 1217018 / 1217018 | 135443 / 135443 | 5072–5128 KiB | 33 / 7419 / 7419 | n/a |
| opaque | family_clone | 3 | 90 | 8946 | 35140 | 26949 | 40 / 40 / 40 | 0 / 0 | 0 / 0 | 4880–5124 KiB | 0 / 0 / 0 | yes |
| opaque | noop | 3 | 90 | 8946 | 35140 | 26949 | 2685740 / 2735810 / 2837990 | 1519140 / 1519140 | 134568 / 134568 | 4944–5100 KiB | 0 / 0 / 0 | yes |
| opaque | add | 3 | 90 | 8946 | 35140 | 26949 | 5297480 / 5378550 / 7036367 | 3200835 / 3200835 | 213359 / 213359 | 4816–4948 KiB | 0 / 0 / 0 | n/a |
| opaque | update | 3 | 90 | 8946 | 35140 | 26949 | 8441723 / 8585692 / 8669853 | 4963609 / 4963609 | 317276 / 317276 | 5204–5356 KiB | 0 / 0 / 0 | n/a |
| opaque | remove | 3 | 90 | 8946 | 35140 | 26949 | 5338430 / 5529341 / 5563351 | 3094087 / 3094087 | 181628 / 181628 | 4892–5072 KiB | 0 / 0 / 0 | n/a |
| opaque | base_edit | 3 | 90 | 8946 | 35140 | 26949 | 7394358 / 7490328 / 7593298 | 4005727 / 4005727 | 317278 / 317278 | 5248–5356 KiB | 0 / 0 / 0 | n/a |

All lanes passed fixture shape, semantic, preservation, changed-edit, inverse, allocator-accounting, and source-observation gates. The `family_clone` and `noop` lanes additionally require the observed Family source pointer to be shared; the gate does not generalize to unrelated owners.

The `native` fixture is `test-data/ooxml/xlsb/date.xlsb`, a small native workbook containing the Office `themeFamily` fragment. The `opaque` fixture is generated in memory from that same package: it retains the native family attributes and adds a family-namespace `extLst` with 96 valid DrawingML `a:ext` children, comments, vendor attributes/payload, and one distinct foreign namespace declaration plus nested payload per child. Its family shape is independently checked before timing.

The synthetic opaque case exercises source-preserving edits at a larger fragment size. It is not a corpus-wide representativeness claim. No broad speedup or regression claim is made from these absolute lanes; a comparison to existing behavior would require a separately frozen control build using the same harness and inputs.

Exact commands, toolchain/host data, binary SHA-256, pre/post Cargo package source manifests, dirty source hashes, raw per-process JSON, and `/usr/bin/time` output are retained beside this report.
Final provenance identifiers: binary SHA-256 `2cae1b47661039b8b382bec814e851836429ee954e32d0cd8f90e780dcb94e69`; matching source manifest SHA-256 `5afedc6a14bc1f363526640a242f33c453103412bc968437b55868e18a03ddd7`.
