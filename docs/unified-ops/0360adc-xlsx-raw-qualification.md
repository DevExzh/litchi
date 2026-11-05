# 0360adc — XLSX raw identity qualification, historical lane

This is a provisional historical result only. The pinned 0bef baseline passed
the raw lifecycle fidelity and allocator checks, and a quiet native Callgrind
run completed. The final current assembly has not been qualified: the OPC
stream changes are still pending and root must supply the final candidate
revision before any comparison or performance conclusion is accepted.

## Reproduced source

- Baseline format source: `0befb112d96d2409e6f40c9abfeb4813665905ad`.
- Overlay assembly snapshot used for the standalone manifest and lock:
  `7858571ad0f4afcb281d54eca313531ef39c9873`.
- Cross-revision probe correction (borrow the relationship target for the
  old `get_part` signature): `b2c7b492ea`; resulting probe SHA-256:
  `a251f1199b5ba442dfd9664806b81f29c7832d071df66fc2d2d5f2ada6c97fd4`.
- Native wrapper commit: `bf31342f35`.
- Standalone manifest SHA-256 after the native wrapper:
  `30ff248844c9e21436fc3913baa4c4d363daffdc617e5e7d03404423410d5c70`.
- Allocation module SHA-256:
  `0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5`.
- Committed lock (`0e24ccbe5c`) SHA-256:
  `8b61b3c5e17b1c32d89139c97a26145117794afae83dd8fef3c9a7aff8765ee6`.

The baseline was built from a disposable archive of 0bef with the standalone
manifest, lock, allocation module, raw probe, and native wrapper overlaid from
the current source. It did not build the full `tools/perf-baseline` package,
so unrelated DOCX macro failures were excluded. The legacy constructor was
selected with `--no-default-features --features allocator-metrics,legacy-xlsx-raw-opc`.

## Evidence

Evidence is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/evidence/raw-qualification-0bef/`.
The protected target is
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63/`;
its pre-existing `xlsx_source_reader_paths_alloc` binary was not overwritten.
The new binaries are named `xlsx-raw-identity-alloc` and
`xlsx-raw-identity-native`.

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| allocator binary | 5,389,480 | `d6563ef99fcd6733e8aab248c578efde67da34fb6ca001cfb3974be878c1565c` |
| native binary | 5,382,704 | `459f0b8b73be43dfcc0e93c7d4fa88a99fe410ca206c5287a7b7aa74c5cba914` |
| allocator/fidelity JSON | 96,846 | `71880f6765ef9be7da2e80f2ff995fe59f930935c9f4b21db0ce2cecfa56238f` |
| native JSON | 40,264 | `93aa2910d2647279335176afb7d1b6ce1430410393d74614a94645aae3bed4d2` |
| Callgrind capture | 392,143 | `3622b535a789054ecbcb9337b39bedb08690bf11cca87b27e18b29dbc176b469` |
| Callgrind JSON | 26,544 | `202a42b4994f314e3ee9c387b87879bb31b6c37e24025f5d238c063a799a96a8` |

The allocator JSON covers 16 operations (four corpora × four operations),
with all allocation samples `measured`, one native identity per operation,
and every untouched-worksheet, replay, inverse, stale-patch, and no-op
boolean true. The native JSON has the same 16/16 fidelity result and reports
allocation status `unavailable` by design. The quiet native Callgrind run used
one warmup and one measured iteration per case and recorded
`3,589,721,324` instructions (`Ir`).

No candidate result, ratio, or acceptance decision is recorded here.
