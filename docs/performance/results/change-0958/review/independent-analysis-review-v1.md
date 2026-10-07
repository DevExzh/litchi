# Independent 0958 allocation-analysis review

Reviewed 2026-10-07 from the sealed six-child packet. This review is read-only and
reports offline arithmetic and source-bound comparisons; it launches no workload,
profiler, or build.

## Custody

The capture packet is `/home/zhuhe/litchi-performance-0958/capture-manifest-v1.json`
(SHA-256
`bd30035f2c7926ec579fea58ecd1e5c5fe34da6b2ae129473bee033630e5c465`). It records
57 files and 92,466,528 bytes. The captured executable is
`/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2`
(SHA-256
`5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3`).

The authoritative offline attribution output reviewed here is
`/home/zhuhe/litchi-performance-0958/analysis-v1/attribute-v3.json`, status
`pass`, SHA-256
`9fb5ae3b1b484e2fc8294c66d4712d8f818c1aed60a974fbe51bcd03173cf983`.
Its input manifest is
`analysis-input-manifest-v1.json`, SHA-256
`0dc7a9a4757d73a05d9fc289581a8ec6add494565f039b94e7addf8d18629495`.
The analyzer used for the final output is
`tools/attribute_0958_v1.py`, SHA-256
`1bb7584076762eefadc409016abf50684157e9350a76386255c5763e47d7cffd`.
The decoder terminal is
`raw-decoder-v1/decoder-terminal.json`, SHA-256
`1bc079e204eda23ab3d8e6c5190760f565fd8d4639d03ac4230e8866b78e6f8b`.

The earlier `attribute-v1.json` is a preserved failed attempt: its folded parser
accepted only tab-separated rows while heaptrack emitted whitespace-separated
stack/count rows. `attribute-v2.json` is also preserved but superseded by v3's
corrected source classification. Any final custody receipt must bind the v3
output to the exact analyzer hash above rather than to an earlier tooling receipt.

The source census used for the ownership comparison is
`review/number-ownership-trace-v2.md`, SHA-256
`3de25114247d354cf6b512db9599aac5ce921149a06fb13351dd048382cd75d2`.
The 0957 allocator phase reference is
`/home/zhuhe/litchi-performance-0957/allocation-summary-replay-v1.json`,
SHA-256
`bc311f3209f346baacedd591a17d6fb36937dda81a256e1ea1cb4f2aaabd4850`.

## Six-run conservation

For every child, the raw trace totals equal both the decoder summary and the
heaptrack size histogram. The selected owner plus unqualified partition equals
the raw total, the two stage-resolution totals equal the selected owner, and
the exclusive leaf partition equals the selected owner. The table gives calls
and requested bytes in each pair.

| child | raw / decoder / histogram | selected owner | unqualified |
| --- | ---: | ---: | ---: |
| h00 medium one-percent r1 | 1,858,480 / 549,846,107 | 330,372 / 51,500,040 | 1,528,108 / 498,346,067 |
| h01 dense one-percent r1 | 3,475,099 / 704,324,249 | 628,254 / 73,812,219 | 2,846,845 / 630,512,030 |
| h02 medium one-edit r1 | 1,182,718 / 448,880,297 | 81,981 / 11,847,945 | 1,100,737 / 437,032,352 |
| h03 medium one-edit r2 | 1,182,722 / 448,880,388 | 81,981 / 11,847,945 | 1,100,741 / 437,032,443 |
| h04 dense one-percent r2 | 3,475,099 / 704,324,238 | 628,254 / 73,812,219 | 2,846,845 / 630,512,019 |
| h05 medium one-percent r2 | 1,858,481 / 549,846,130 | 330,372 / 51,500,040 | 1,528,109 / 498,346,090 |

Each run retains one unknown-any-frame and unknown-leaf-frame row of 73,728
bytes in the unqualified partition. Missing-owner reason counters can overlap
and must not be summed as disjoint rows. Folded call totals also match, but
heaptrack folded output has no requested-byte column, so folded byte conservation
is explicitly unavailable rather than inferred.

The aggregate is 13,032,599 raw calls / 3,406,101,409 requested bytes and
2,081,214 selected calls / 274,320,408 requested bytes.

## Stage boundary and phase cross-check

The selected owner requires `timed_iteration` plus either the symbolized
`stage_timed_commit` owner or the reviewed `MultiSourceEdit::commit` owner, and
rejects rows with unknown frames. The two resolution buckets are disjoint in the
output and sum exactly to the selected total:

| workload | `stage_timed_commit` | `MultiSourceEdit::commit` | selected total |
| --- | ---: | ---: | ---: |
| medium one-percent (each repeat) | 204,636 / 42,124,623 | 125,736 / 9,375,417 | 330,372 / 51,500,040 |
| dense one-percent (each repeat) | 390,342 / 56,308,476 | 237,912 / 17,503,743 | 628,254 / 73,812,219 |
| medium one-edit (each repeat) | 51,456 / 9,674,712 | 30,525 / 2,173,233 | 81,981 / 11,847,945 |

The exact phase-counter cross-check is against three worker-one 0957 samples:

- Medium one-percent `edit_commit`: 110,124 calls / 17,166,680 bytes;
  multiplied by three, 330,372 / 51,500,040.
- Dense one-percent `edit_commit`: 209,418 / 24,604,073; multiplied by three,
  628,254 / 73,812,219.
- Medium one-edit `edit_commit`: 27,327 / 3,949,315; multiplied by three,
  81,981 / 11,847,945.

This is a strong counter-level corroboration of the phase boundary. It is not a
claim that every selected stack is independently a complete `edit_commit`
interval: `stage_timed_commit` is the full stage callsite when symbolized,
whereas `MultiSourceEdit::commit` is a reviewed commit-core surrogate and does
not by itself imply the complete edit-commit phase. The combined selected total
is the admitted phase-equivalent cross-check, not a new allocation inference.

## Source ownership cross-check

The source census reports omitted numeric owners and lexical bytes per measured
iteration. The nonadditive `stored_clone` intersection in each heaptrack run
matches that census multiplied by three:

| workload | source census per iteration | three-iteration `stored_clone` result |
| --- | ---: | ---: |
| medium one-percent | 9,123 / 58,723 | 27,369 / 176,169 |
| dense one-percent | 17,614 / 92,907 | 52,842 / 278,721 |
| medium one-edit | 2,303 / 10,933 | 6,909 / 32,799 |

`number`, `clone`, `merge`, `number_clone`, `merge_clone`, and related source
tags are overlapping diagnostics. They are not additive and must not be added
to `stored_clone` or to the exclusive partition. The v3 exclusive buckets are
separately checked to sum exactly to the selected owner.

These results establish the bounded call/byte correspondence and preserve all
unknown and unqualified work. They do not establish causal timing, OS behavior,
physical I/O, or a production allocation saving.
