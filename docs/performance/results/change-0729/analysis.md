# Public DOC attribution

All ranges span nine independent process statistics per case/route. Whole timing is microseconds. Fractions are computed within each measured owner and then summarized per process. No cross-route median is subtracted to invent a nested phase.

| Case | Route | Whole p50 μs | Whole mean μs |
| --- | --- | ---: | ---: |
| docfloat | ordinary-opaque | 990.32–1135.86 | 994.93–1128.18 |
| docfloat | ordinary-split | 991.65–1145.82 | 993.17–1142.06 |
| docfloat | profiled-empty | 1001.26–1105.11 | 1016.98–1086.42 |
| docfloat | profiled-clock | 1021.09–1094.58 | 1034.35–1085.35 |
| docnohf | ordinary-opaque | 105.74–107.32 | 106.93–109.54 |
| docnohf | ordinary-split | 105.51–108.17 | 106.58–109.93 |
| docnohf | profiled-empty | 105.42–108.80 | 107.05–109.93 |
| docnohf | profiled-clock | 107.26–109.46 | 108.78–111.52 |

## Outer phases, ordinary split

| Case | Phase | Median % of same owner whole |
| --- | --- | ---: |
| docfloat | open_ns | 25.95–29.22 |
| docfloat | edit_ns | 7.72–8.83 |
| docfloat | replace_ns | 19.97–22.45 |
| docfloat | commit_ns | 37.72–42.82 |
| docfloat | output_copy_ns | 0.39–0.57 |
| docfloat | whole_residual_ns | 0.06–2.44 |
| docnohf | open_ns | 22.17–22.51 |
| docnohf | edit_ns | 7.69–7.88 |
| docnohf | replace_ns | 32.15–32.78 |
| docnohf | commit_ns | 36.40–36.77 |
| docnohf | output_copy_ns | 0.36–0.43 |
| docnohf | whole_residual_ns | 0.25–0.29 |

## Semantic phases, profiled clock

| Case | Phase | Median % of same owner whole | Median % of parent phase |
| --- | --- | ---: | ---: |
| docfloat | open.StrictOwnerValidation | 9.69–10.19 | 34.12–34.70 |
| docfloat | open.PublicReaderValidation | 17.17–18.29 | 61.20–61.85 |
| docfloat | open.SourceRetention | 0.54–0.57 | 1.91–1.96 |
| docfloat | commit.Finish | 7.68–8.18 | 18.97–21.58 |
| docfloat | commit.StrictOwnerValidation | 9.89–13.67 | 25.47–33.73 |
| docfloat | commit.PublicReaderValidation | 17.20–18.13 | 42.43–48.31 |
| docfloat | commit.SourceRetention | 0.58–3.46 | 1.46–8.84 |
| docfloat | commit.Patch | 0.01–0.01 | 0.02–0.03 |
| docnohf | open.StrictOwnerValidation | 9.66–9.76 | 41.35–41.78 |
| docnohf | open.PublicReaderValidation | 12.33–12.45 | 52.82–53.23 |
| docnohf | open.SourceRetention | 0.42–0.46 | 1.80–2.01 |
| docnohf | commit.Finish | 13.00–13.30 | 37.06–37.60 |
| docnohf | commit.StrictOwnerValidation | 9.42–9.57 | 26.75–27.17 |
| docnohf | commit.PublicReaderValidation | 11.36–11.62 | 32.14–32.73 |
| docnohf | commit.SourceRetention | 0.53–0.55 | 1.48–1.56 |
| docnohf | commit.Patch | 0.05–0.06 | 0.13–0.18 |

## Matched observer controls

Positive deltas are slower. Flags count absolute changes above 5% in either direction; they are interpretation flags, not optimization admission gates. Each pair matches case, cycle and round.

| Case | Before → after | p50 delta % | p50 flags / 9 | Mean delta % | Mean flags / 9 |
| --- | --- | ---: | ---: | ---: | ---: |
| docfloat | ordinary-opaque → ordinary-split | -9.75–15.70 | 4 | -6.13–14.20 | 5 |
| docfloat | ordinary-split → profiled-empty | -9.24–7.58 | 4 | -7.89–2.40 | 2 |
| docfloat | profiled-empty → profiled-clock | -5.63–5.44 | 2 | -3.45–4.85 | 0 |
| docnohf | ordinary-opaque → ordinary-split | -1.68–2.07 | 0 | -1.24–1.89 | 0 |
| docnohf | ordinary-split → profiled-empty | -2.35–2.90 | 0 | -2.06–2.62 | 0 |
| docnohf | profiled-empty → profiled-clock | -0.23–2.99 | 0 | -1.05–2.20 | 0 |

Profiled open releases its strict editor before subsequent validation, unlike ordinary open. The profiled implementation comparison therefore includes lifetime/code-path changes. Timestamp recorder calibration is measured outside the workflow and is never subtracted from workflow timing. This build enables performance-diagnostics for every route; the experiment does not measure feature-enabled versus default-build code generation.
