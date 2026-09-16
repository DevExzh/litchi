# Evidence packet: change 0637

Parse the eager PPTX slide catalog once per borrowed `Presentation` instead of
once per catalog query, and measure PPTX-4 — the three passes
`semantic_text_from_part` runs per slide — without implementing it.

Record: [`../../0637-pptx-eager-slide-catalog-memo.md`](../../0637-pptx-eager-slide-catalog-memo.md).
Implements **PPTX-3** of [0587](../../0587-remaining-opportunity-survey.md) and
freezes the design for **PPTX-4 / DOCX-3**.

## Provenance

| field | value |
| --- | --- |
| base commit (before leg) | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` |
| branch | `perf/0637-pptx-eager-slide-catalog-memo` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0, valgrind 3.26.0 |
| build | `cargo build --release --locked` for the harness; the scratch probe adds `debug = 1` to the same `lto = true`, `panic = "abort"` profile |
| pinning | every measured process `taskset -c 13` |
| host state | seven other agents were building and measuring concurrently; run-window load average 15-33 (`timing/window.txt`) |

Binary SHA-256:

| binary | sha256 |
| --- | --- |
| before `litchi-perf-baseline` | `2d7421fad0cee46ed59c3dc61895ac7db64d710f240716aafc64e86fb4baa43e` |
| after `litchi-perf-baseline` | `5abbc4669a1f9faaa394fd82b8409b57302cb5c471204d79acf162b2c0d0506c` |
| before `probe0637-pptx` | `9dd79eb475e2beae216581290a4fde15cfc42aa73948b5f5815f85d67861ac91` |
| after `probe0637-pptx` | `dc271313f26410d06ae33426ae74816e1cf79cc5d9487d81e40a560659425195` |
| PPTX-4 prototype `probe0637-pptx` | `66efd60a53790db9a25c56759ae15ffd4c1009dee8562ff41ac469a38428faff` |

All four timed binaries were staged outside any Cargo target directory before
being run (change 0627's lesson: a concurrent build relinked one mid-run).

## Corpora

| deck | slides | bytes | where it comes from |
| --- | ---: | ---: | --- |
| `test-data/ooxml/pptx/shapes.pptx` | 6 | 68,822 | repository fixture |
| `test-data/poi/test-data/slideshow/bug62513.pptx` | 19 | — | repository fixture, the largest slide text of the 78 |
| `test-data/poi/test-data/slideshow/45545_Comment.pptx` | 11 | — | repository fixture |
| harness 200-slide deck | 200 | 17,017,139 | `tools/perf-baseline`'s `build_pptx_source_edit_corpus`, 8 text boxes per slide and 8 × 2 MiB PNGs, sha256 `61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757` |

The 200-slide deck is not retained here — it is 17 MB and reproducible.
`fixtures/capture-harness-deck.sh` regenerates it; the harness materializes it
under `--filesystem-root` and deletes the run directory afterwards, and exposes
no corpus-export command, so the script copies it out while a long run holds it.

## Contents

| path | what it is |
| --- | --- |
| `probe/src/main.rs` | the scratch probe: `counts` and `time` for eleven catalog and text operations `tools/perf-baseline` has no selector for, `oracle` for the cross-leg corpus census, `build` for an authored deck |
| `probe/Cargo.toml.example` | its manifest template; point the path dependencies at the leg being measured |
| `counts/run-counts.sh` | the callgrind isolation-pair runner for the catalog operations, both legs |
| `counts/run-pptx4.sh` | the same for PPTX-4, before leg against the raw-scan-free prototype |
| `counts/isolation-summary.py` | differences each pair into per-operation Ir |
| `counts/callgrind-raw-totals.txt` | the raw `Collected` totals behind the catalog pairs |
| `counts/callgrind-isolation-summary.txt` | per-operation Ir, both legs, and the delta |
| `counts/pptx4-raw-totals.txt`, `counts/pptx4-isolation-summary.txt` | the same for PPTX-4 |
| `counts/pptx4-symbol-attribution.txt` | `callgrind_annotate --inclusive` for `semantic_text_from_part`, `process_markup_compatibility` and the parser on the 200-slide deck, both legs — the per-pass split in the record |
| `timing/timing-harness.sh` | the paired ABBA runner for the three named selectors |
| `timing/timing-harness-repeat.sh` | the confirmation window for the two eager filesystem selectors, with two unreachable controls |
| `timing/timing-probe.sh` | the paired ABBA runner for the probe operations and for PPTX-4 |
| `timing/stats.py` | percentiles, paired deltas in both directions, and the A/A and B/B floors |
| `timing/harness-samples.py` | extracts per-sample nanoseconds for one case and shape from a harness report |
| `timing/catalog-*/`, `timing/pptx4-*/` | per-sample nanoseconds per operation, A1 B1 B2 A2, 40 samples per leg-run |
| `timing/harness/*/` | the same, split out of the harness reports |
| `timing/harness-repeat/*/` | the confirmation window's samples |
| `timing/probe-timing-summary.txt`, `timing/harness-timing-summary.txt`, `timing/harness-repeat-summary.txt` | the computed percentiles and floors |
| `timing/window.txt` | wall-clock start and end of each timing window with the host load average |
| `oracle/pptx-corpus-before.txt`, `oracle/pptx-corpus-after.txt` | 78 `.pptx` fixtures × every catalog-dependent public projection = 1,229 rows per leg, byte-identical |
| `reach/route-census.txt` | which `litchi-pptx` presentation route the two eager filesystem selectors actually take |
| `pptx4-raw-scan-removed-prototype.patch` | the measurement-only edit behind the PPTX-4 upper bound. **Not a legal production change**: it deletes an adversarial-input defence |
| `fixtures/capture-harness-deck.sh` | regenerates the 200-slide deck |
| `decision.json` | the machine-readable decision record |
| `gates.txt` | the tail of every gate |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |

## Replay

```sh
# one probe per leg
cp -r probe /tmp/probe-after && cd /tmp/probe-after
sed -i "s|REPLACE_WITH_CHECKOUT|<checkout>|" Cargo.toml.example && mv Cargo.toml.example Cargo.toml
CARGO_TARGET_DIR=<target> cargo build --release

# counts (deterministic)
bash counts/run-counts.sh <scratch>   > counts/callgrind-raw-totals.txt
bash counts/run-pptx4.sh  <scratch>   > counts/pptx4-raw-totals.txt
python3 counts/isolation-summary.py counts/callgrind-raw-totals.txt
python3 counts/isolation-summary.py counts/pptx4-raw-totals.txt before,proto

# oracle
<probe> oracle test-data > oracle/pptx-corpus-<leg>.txt

# timing (one window, deterministic counts first)
bash timing/timing-harness.sh <scratch>
bash timing/timing-probe.sh   <scratch>
python3 timing/stats.py timing/catalog-deck200-slideall
```
