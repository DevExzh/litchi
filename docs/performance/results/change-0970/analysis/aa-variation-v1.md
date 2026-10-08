# 0969 DOC A/A variation analysis

This is a descriptive offline analysis of the retained baseline-only A/A raw packets. It launches no workload, build, test, or reducer and makes no host-cause claim.

Custody revalidation: **PASS**. The sealed 0969 audit, plan/GO/terminal/readiness/stage bindings, and raw packet set were rechecked before analysis. The packet set contains 576 children and 46080 measured samples.

Within-child sample CV has median **3.034%** and p95 **6.682%**. Across the 24 cell-level groups, the between-child median CV has median **1.508%**; the between/within CV ratio has median **0.881**. Between-child median dispersion exceeds within-child CV in 10/24 cells and exceeds it 2× in 6/24; this is a descriptive scale comparison, and cell and order composition remain part of the measured grouping.

Across cells, the BA/AB median ratio is 1.0011 (p10–p90 0.9482–1.0184); the second/first position ratio is 0.9996 (p10–p90 0.9739–1.0219). Median chronological Spearman rho is 0.053; median pair-index rho is -0.003. These are mixed cell-level trends, not evidence of a common direction or cause.

Family medians drift -1.449% for docnohf and 0.370% for docfloat. The edit-group medians range -1.857%–-0.943% for docnohf and -0.145%–0.923% for docfloat; retention-group medians range -1.640%–-1.376% and -0.086%–0.730%, respectively. The heuristic candidates are 12 docfloat and 0 docnohf.

Each child declares five warmups, but warmup samples are not retained in the report. The retained measured sequence compares its first and last ten samples: overall median drift is **-0.865%** (p05–p95 -5.933%–3.383%). This describes measured-sequence drift after warmup; it cannot identify warmup behavior itself.

The conservative two-level sequence screen flags 12 of 576 children as heuristic candidates. It requires a 10× pooled-MAD sorted gap, at least a 5% center separation, at least ten observations per side, and at least three chronological runs. The screen is exploratory and does not establish a statistical bimodality test or a host explanation.

Largest heuristic candidates:

- `aa-docfloat-e1-default-p01-first`: low/high medians 966244/1216060 ns, gap/MAD 31.1, runs 4, transitions 3.
- `aa-docfloat-e2-zero-p04-first`: low/high medians 1266960/1480121 ns, gap/MAD 26.0, runs 8, transitions 7.
- `aa-docfloat-e8-zero-p08-second`: low/high medians 2562736/3394159 ns, gap/MAD 16.9, runs 16, transitions 15.
- `aa-docfloat-e1-default-p06-first`: low/high medians 983534/1192035 ns, gap/MAD 16.5, runs 8, transitions 7.
- `aa-docfloat-e8-zero-p06-second`: low/high medians 2592266/3422784 ns, gap/MAD 15.9, runs 4, transitions 3.

The 0969 readiness result remains the governing A/A outcome: complete evidence exceeded its frozen symmetric noise bound. This variation report does not convert that failure into a workload or host diagnosis.

Analyzer: `/home/zhuhe/litchi-performance-0970/tools/analyze_aa_variation_v1.py` (SHA-256 `5543972f07ad845fe22564f1876e13c6fe3388111fe9f3037e738800345660c6`). JSON receipt: `/home/zhuhe/litchi-performance-0970/analysis/aa-variation-v1.json` (SHA-256 `9e41ff784382ae5d225ce751d3398eac1ee1ff56500d6e229d68508ef07f710a`).
