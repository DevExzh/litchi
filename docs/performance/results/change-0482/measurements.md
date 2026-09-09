# XML audit measurements

All values derive from `summary.json`; raw reports remain under `captures/`. Each row has 30 samples after three warmups. Intervals are the recorded t(29) intervals for the sample mean. RSS covers the complete process, including setup and teardown. Normal and allocator timings remain separate.

## Normal timing and process RSS

| Route | Bytes | Repeat | Mean ms [95% interval] | p50 / p95 / p99 ms | RSS MiB |
| --- | ---: | ---: | --- | --- | ---: |
| materialized | 65,536 | 1 | 0.089028 [0.087769, 0.090288] | 0.087761 / 0.099161 / 0.100910 | 2.164062 |
| materialized | 65,536 | 2 | 0.088942 [0.088135, 0.089749] | 0.088051 / 0.095210 / 0.095331 | 2.382812 |
| materialized | 8,388,608 | 1 | 10.996315 [10.993732, 10.998898] | 10.995569 / 11.010418 / 11.012848 | 10.148438 |
| materialized | 8,388,608 | 2 | 10.965352 [10.937444, 10.993260] | 10.943548 / 11.208809 / 11.259519 | 10.160156 |
| materialized | 134,217,728 | 1 | 204.054625 [203.637281, 204.471969] | 204.236211 / 205.569937 / 205.994649 | 130.179688 |
| materialized | 134,217,728 | 2 | 203.475179 [203.163522, 203.786835] | 203.369229 / 205.075876 / 205.644918 | 130.152344 |
| streaming | 65,536 | 1 | 0.088799 [0.087931, 0.089667] | 0.088080 / 0.094690 / 0.098071 | 1.898438 |
| streaming | 65,536 | 2 | 0.089081 [0.088242, 0.089919] | 0.088170 / 0.094930 / 0.095500 | 1.921875 |
| streaming | 8,388,608 | 1 | 11.390709 [11.367497, 11.413922] | 11.375821 / 11.437121 / 11.705542 | 1.921875 |
| streaming | 8,388,608 | 2 | 11.414038 [11.386548, 11.441529] | 11.392550 / 11.671591 / 11.685011 | 1.929688 |
| streaming | 134,217,728 | 1 | 181.998491 [181.900437, 182.096545] | 181.944003 / 182.383694 / 182.909927 | 1.929688 |
| streaming | 134,217,728 | 2 | 181.923757 [181.859386, 181.988128] | 181.887092 / 182.293254 / 182.496675 | 1.902344 |

## Allocator timing and process RSS

| Route | Bytes | Repeat | Mean ms [95% interval] | p50 / p95 / p99 ms | RSS MiB |
| --- | ---: | ---: | --- | --- | ---: |
| materialized | 65,536 | 1 | 0.088484 [0.087799, 0.089169] | 0.087801 / 0.093580 / 0.095461 | 2.382812 |
| materialized | 65,536 | 2 | 0.089104 [0.088057, 0.090151] | 0.088010 / 0.097230 / 0.098650 | 2.167969 |
| materialized | 8,388,608 | 1 | 10.977372 [10.961418, 10.993327] | 10.969818 / 11.050858 / 11.181689 | 10.175781 |
| materialized | 8,388,608 | 2 | 10.964988 [10.960333, 10.969642] | 10.965536 / 10.984616 / 10.987526 | 10.171875 |
| materialized | 134,217,728 | 1 | 202.840903 [202.425972, 203.255834] | 202.519840 / 204.562549 / 205.464583 | 130.179688 |
| materialized | 134,217,728 | 2 | 203.707239 [203.380308, 204.034170] | 203.750822 / 204.808114 / 205.734250 | 130.132812 |
| streaming | 65,536 | 1 | 0.098506 [0.079197, 0.117814] | 0.088491 / 0.096730 / 0.372081 | 1.902344 |
| streaming | 65,536 | 2 | 0.089336 [0.088414, 0.090258] | 0.088390 / 0.094540 / 0.097960 | 1.929688 |
| streaming | 8,388,608 | 1 | 11.408699 [11.388226, 11.429172] | 11.393540 / 11.565471 / 11.615890 | 2.140625 |
| streaming | 8,388,608 | 2 | 11.413006 [11.404244, 11.421769] | 11.407470 / 11.461890 / 11.492740 | 1.921875 |
| streaming | 134,217,728 | 1 | 182.451955 [182.386549, 182.517362] | 182.446463 / 182.810804 / 182.827958 | 1.898438 |
| streaming | 134,217,728 | 2 | 182.423892 [182.350824, 182.496960] | 182.404676 / 182.725598 / 182.917959 | 2.140625 |

## Operation allocation observations

Each value below is identical in all 60 samples for that route and size. Allocation calls include reallocations where present. All 360 allocator samples finish with zero net live bytes and no failed allocation calls.

| Route | Input bytes | Peak increment bytes | Requested bytes | Allocation calls | Reallocation calls |
| --- | ---: | ---: | ---: | ---: | ---: |
| materialized | 65,536 | 147,504 | 278,544 | 17 | 12 |
| materialized | 8,388,608 | 16,793,648 | 33,570,832 | 24 | 19 |
| materialized | 134,217,728 | 268,451,888 | 536,887,312 | 28 | 23 |
| streaming | 65,536 | 65,587 | 65,587 | 7 | 0 |
| streaming | 8,388,608 | 65,587 | 65,587 | 7 | 0 |
| streaming | 134,217,728 | 65,587 | 65,587 | 7 | 0 |

## Corpus identities

The record text byte is `a + ((record_index mod 256) mod 26)`. The source window is 16 KiB; a regular item record is 1,024 bytes. The final item absorbs the remaining body bytes. Each raw report binds its independently checked generated-byte digest.

| Input bytes | SHA-256 of generated XML |
| ---: | --- |
| 65,536 | `2a01e49ced5867dee167dfbe60b795604e00e7ed094adc2aa805636a81e0aa1f` |
| 8,388,608 | `d29f983b3678a7af5fcd03504676d1652d091695bc2ea5f1e3d3bae3001e4d8b` |
| 134,217,728 | `834a81788afe1eca89aeef88e943c320e20be43bb9e93b46c012f792d3d4ec0b` |

## Review triggers

The analyzer flags absolute changes over 5%, including improvements: 17 pair/repeat comparisons are flagged. Adverse observations are the allocator 64 KiB streaming R1 mean (+11.326% versus materialized), normal 64 KiB materialized repeat RSS (+10.108%), and allocator 128 MiB streaming repeat RSS (+12.757%). The allocator 64 KiB streaming R1 data include a 372,081 ns sample; it remains in the reported mean and tails. Its R2 mean is 9.308% lower than R1.

At 8 MiB, normal streaming means are 3.587% and 4.092% higher than materialized; both observations remain visible despite falling below the 5% trigger. At 128 MiB they are 10.809% and 10.592% lower. No normal latency repeat change crosses 5%.

The large heap difference includes the materialized route’s `read_to_end` capacity growth. It is not a DOCX transaction comparison, a constant-RSS claim, or an end-to-end package-memory guarantee.
