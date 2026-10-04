# 0902 independent raw result review

Status: **PASS**. Decision: **rejected; no production adoption**.

This review reads the retained JSON reports and recomputes the native paired
ratios and bootstrap intervals offline. It did not invoke `run.py`,
`analyze_v3.py`, `replay.py`, or any workload/build command.

## Bound evidence

| Artifact | SHA-256 |
| --- | --- |
| `analysis.json` | `9ab725ba8790245fe9f7cca0adaf1f21d458c0e715f9aceed17ea1cac52e14bc` |
| `native/capture.json` | `5cb36d9ccf5061ac48cfa629fd67b638e4ce5903fd20280ea997bbba57177d72` |
| `allocator/capture.json` | `eca6ea0197f27e6414f6960f5e79c0f41541185bf99d7044ab4b49105782ad6e` |
| `raw-replay.json` | `1e06ac15760ef365cd20b9cc0a5eea081097f98e823a2df292d00c1dbbf88a3d` |
| `decision.json` | `8d79ae722b8576b1567ca6720007484d09031b0f9a432eaefae31e6ed17bc037` |
| `source-final.json` | `d4aabd8c1fe6dc44d961e9e36c3aa1857f49faf6cf447897a07cccf481edfd51` |
| candidate `audit.rs` overlay | `e1d73295becd05347dfa6c07910089f16e60b44a6eeb50d92938a3a50cb10fdb` |
| candidate `source_replacement.rs` overlay | `f7f05d3e2f733c1586c9402f79b29ba4087c8148b7afc88bc8a2c86c6e246d20` |

The retained replay reports 10,771 verified baseline files, the two candidate
override files, and 8 qualification, 48 native, and 16 allocator children.
The complete per-child report descriptors and allocator phase-field records
are in `result-review.json`.

## Native rows and gate

The raw 48 native reports contain six blocks for each of four cases, with 100
samples and 10 warmups per child. The checker independently recomputed the
integer p50, nearest-rank p95/p99, sample ordering, operation-metric alignment,
phase alignment, and sink limits. The table lists each block as
`baseline/candidate (ratio)`.

| Case | Six block p50 pairs | Paired p50 median | 95% bootstrap interval | Primary gate |
| --- | --- | ---: | ---: | --- |
| `primary-percent-medium` | 16703700/16638275 (0.996083203), 16628545/16730031 (1.006103120), 16621451/16733042 (1.006713674), 16697087/16690741 (0.999619934), 16618036/16643462 (1.001530024), 16525516/16612323 (1.005252907) | 1.0033914657 | [0.9978515684, 1.0064083967] | FAIL |
| `primary-percent-dense-sparse` | 32959875/34078575 (1.033941269), 33511028/34208883 (1.020824637), 32997243/32521039 (0.985568370), 33261357/32840245 (0.987339302), 32844888/33007742 (1.004958275), 32782808/32793452 (1.000324682) | 1.0026414789 | [0.9864538360, 1.0273829533] | FAIL |
| `control-xlsx-one-edit-medium` | 4147051/4173356 (1.006343062), 4155711/4088250 (0.983766677), 4152686/4224587 (1.017314336), 4134506/4108217 (0.993641562), 4095291/4108212 (1.003155087), 4082466/4120807 (1.009391628) | 1.0047490744 | [0.9887041194, 1.0133529817] | N/A (control) |
| `control-docx-one-edit-medium` | 166641/157116 (0.942841198), 164936/158926 (0.963561624), 165721/157420 (0.949909788), 165281/157536 (0.953140409), 165541/158326 (0.956415631), 168786/157831 (0.935095328) | 0.9515250988 | [0.9389682627, 0.9599886275] | N/A (control) |

The bootstrap used 10,000 seeded draws (`902009`) with nearest-rank indices 249
and 9749. Both primary cases fail the required p50 median ratio <= 0.98:
medium is 1.0033914657, and dense-sparse is 1.0026414789.
The recomputed 95% intervals, block ratios, and empty review-flag list match
`analysis.json` exactly. The final primary gate is `false`.

## Semantic and allocator checks

All 72 retained rows normalize to one semantic identity per case across
qualification, native, and allocator populations. XLSX source/output/semantic
digests, cache counters, source limits, phase alignment, and compressed-read
classification were checked. The DOCX control retains no fabricated XLSX
source or output digest.

All 16 allocator reports use
`system_allocator_operation_scoped` / `serialized_region_peak_v3` and contain
three measured samples. The 12 XLSX allocator reports expose five measured
phase allocation vectors each; the four DOCX control reports intentionally
contain no phase metrics. The review checked 60 phase vectors and 180
phase-sample metric records, zero failed allocation calls, nonnegative
counters, live/peak boundary relationships, operation-metric alignment, and
the 65,536-byte sink limit. No allocator boundary violation was found;
allocator evidence remains diagnostic and is not adopted.

`source-final.json` was checked while the candidate overlay was live and records
that overlay as evidence. `decision.json` records the subsequent restoration
of the two production source files to their baseline hashes, along with
`rejected`, `production_retained: false`, and no adopted improvement.
