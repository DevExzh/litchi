# 0892 — current producer XLSX full-cell profile

This batch refreshes hotspot evidence after the adopted [0891 x14ac change](0891-x14ac-namespace-resolution.md). MCE remains the largest named parser group: **28.65–29.70%** of whole-process sampled user-cycle periods. The separate runner-scoped guest diagnostic attributes **46.46%** of instructions inclusively to MCE processing. These are different populations and scopes; neither establishes a new latency result.

Production and benchmark source are unchanged. The next candidate should remove repeated full worksheet parsing through a conservatively qualified selected route, after strengthening the benchmark's per-sheet tolerance checks.

## Source and protocol

The [source receipt](results/change-0892/source-identity.json) verifies 10,765 live selected files against the 0885 source archive plus the two sealed 0891 overlays. Four separate `xml_attribute_bounds` diagnostic files are accounted for outside the profiled executable. The unrelated untracked `bench_word_parse.rs` example is untouched.

The reused Rust 1.95 release executable is the adopted 0891 native binary, SHA-256 `1cc6b194f4a694eaaddedee95c5dcf4e472e1790306926cf6cb9b1aff378c710`, ELF build ID `22608b8f45f6d0d2a8e2a71dc14435bd67f8a320`, at source revision `0a0957ef7249b4b16ec26ed81d1fb32b7288bb0b`. No rebuild was necessary. The [methodology review](results/change-0892/methodology-review.json) and two frozen plans bind source, executable, tools, helpers, and prior comparison inputs.

The case is `xlsx_producer_medium_source_full_cell_scan`: a generated 22,739-byte archive, four sheets, 4,096 cells, 3,686 numeric values and 410 shared-string values. All five child reports replay the archive, source fence, cell digest, census, and complete semantic vectors. The existing producer fallback materializes its stores, so its aggregate tolerance guard remains valid for this batch.

Both populations use sequential fresh children, CPU 12, warm filesystem cache, and the shared host. Three native `perf` captures use `cycles:u`, 499 Hz, `dwarf,16384`, ten warmups and 2,000 measured scans each. Two separate Callgrind captures toggle collection at `run_xlsx_full_cell_scan`, with no warmups and ten scans each. There were no retries or replaced children. Instrumented elapsed times are retained but are not latency comparisons.

## Native sampled-IP observations

These disjoint groups count sampled instruction-pointer self periods over the whole process, including setup, timer-excluded checks, destruction, and reporting. Ranges describe three children in each historical population; they are not paired confidence intervals.

| Self-symbol group | 0886 period share | Current period share |
| --- | ---: | ---: |
| MCE | 26.73–27.78% | 28.65–29.70% |
| x14ac | 5.69–6.36% | 4.54–5.11% |
| XML namespace resolver | 6.87–10.16% | 2.60–2.98% |
| XML namespace reader | 4.00–5.13% | 3.64–4.34% |
| Other worksheet parser | 9.62–10.21% | 10.24–11.16% |
| Other quick-xml | 9.49–10.62% | 9.93–10.94% |
| Unresolved sampled IP | 0.60–0.92% | 0.78–0.98% |

The current children contain 3,817 / 3,814 / 3,800 samples and exactly 31,519,530,592 / 31,665,956,013 / 31,453,030,639 periods. Flat-symbol and leaf-sample totals reconcile; all report zero lost samples. Empty-stack period shares are 83.96% / 0.00% / 78.22%, and shares with an unknown stack frame are 3.42% / 45.73% / 6.14%. The second child's nonempty stacks therefore do not establish reliable caller attribution. No native inclusive attribution is used.

The smaller namespace-resolver share is consistent with 0891's eliminated lookups. Shares depend on the denominator and cannot independently prove a speedup or an MCE regression. Full disjoint groups and top symbols are in [self-analysis.json](results/change-0892/self-analysis.json).

## Separate guest-instruction observations

| Inclusive canonical row | 0886 runner share | Current runner share |
| --- | ---: | ---: |
| MCE process | 44.10–44.19% | 46.46–46.46% |
| MCE start | 28.81–28.85% | 30.34–30.35% |
| x14ac capture | 19.21–19.24% | 15.00–15.00% |

Current runner totals are 717,244,483 and 717,348,440 guest instructions, versus 755,248,408 and 754,020,187 in 0886. Current x14ac inclusive counts are 107,567,302 and 107,567,742, versus 145,048,571 and 145,049,900. MCE processing remains approximately 333 million instructions per child.

The runner includes oracle preparation and timer-excluded work. MCE start is nested within MCE process; inclusive rows overlap and must not be summed. Both Valgrind `brk segment overflow` warnings are retained. Guest instruction counts are not native cycles, wall-clock latency, allocation measurements, or speedup evidence. Canonical selectors, parent PIDs, exact totals, commands and comparisons are in [callgrind-analysis.json](results/change-0892/callgrind-analysis.json).

## Next work and required qualification

The [static route audit](results/change-0892/streaming-route-audit.md) predicts the first producer refusal at `<sheetViews>` as `UnsupportedStructure`, before `<cols>`. The existing `selected_cell_stream_ineligible_on_cols` field only records byte presence; it is not a runtime reason census. Later barriers include sheet format, columns, row spans and page metadata.

A narrow eligible metadata subset could avoid the fallback's second full XML/model parse and Store construction while retaining the selected MCE+x14ac stream, EOF verification, physical record retention, dependency checks, source fences and callback ordering. No gain is predicted before measurement.

Before implementing that candidate, add actual refusal-reason evidence and force every worksheet's tolerance report outside the timer. Workbook-level reports inspect only published stores, so they could miss repairs after a selected route leaves a store unmaterialized. Repairable, unknown, styled, duplicate or misordered metadata must fall back to the authoritative parser. Qualification must cover typed errors, limits, cancellation, source changes, shared strings, relationships, MCE/x14ac handling, and OPC CRC/size failures before any callback is exposed.

## Validation and retained evidence

[Packet verification](results/change-0892/verification.json) passes all five children / 6,020 measured scans, command and artifact custody, source and executable identity, ELF build IDs, semantic replay, exact native period reconciliation, canonical guest rows and historical comparison recomputation. Helper AST and retained-0886 compatibility checks also pass. No Rust source changed, so this diagnostic batch did not rebuild or repeat the already passing 0891 owner quality suite.

Independent review and post-cleanup replay are retained alongside the packet. No worktree, source checkout or build target was created for 0892. Raw perf data and the reused executable remain locally retained and ignored; text exports, Callgrind raw data and receipts are committed. Existing unrelated worktrees remain untouched. The [packet README](results/change-0892/README.md) gives replay commands and limitations.

This batch completes current-source diagnosis only. The broader non-iWork performance goal remains open; 0891 remains the scoped latency adoption evidence.
