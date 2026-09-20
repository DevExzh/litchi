# Independent read-only distribution review

All four failed comparisons retain all 100 native samples. Each is mean-only;
the corresponding p50 passes. The following values were independently checked
against analysis.json and measurement-details.json.

| Case / mode / metric | Pair | Mean delta | p50 delta | p95 delta | p99 delta | Maximum baseline → candidate (ns) |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| 54016 stored / owned / q3 | B2/A2 | +8.679% | 0.000% | +1.754% | +2.586% | 1,590 → 7,820 |
| 54016 stored / file / q3 | B1/A1 | +5.587% | −0.194% | +0.375% | +240.727% | 2,820 → 11,230 |
| Generated / owned / q3-to-q8 mean | B2/A2 | +5.876% | +2.102% | +2.549% | +41.361% | 865 → 2,346.67 |
| 45365 late / file / q8 | B2/A2 | +8.399% | +1.744% | +2.273% | +1.111% | 1,800 → 12,750 |

For those rows in order, A/A mean deltas are −1.057%, +0.654%, −4.278%,
and +0.647%. A2/A1 mean deltas are +2.213%, +2.435%, +1.344%, and −1.580%;
B2/B1 mean deltas are +11.151%, −2.353%, +5.193%, and +7.746%.
Several failed B2/A2 means coincide with larger candidate maxima, while the
failed B1/A1 row has its larger observations in B1. The pattern does not
establish either a systematic candidate regression or a scheduling/hardware
cause. The frozen mean gate still fails. No samples were excluded, resampled,
or reweighted for this analysis.

The complete matrix retains 119 paired upper-tail flags, 14 A/A central drift
flags and 28 within-phase central drift flags. These diagnostics explain
sensitivity to upper-tail observations; they do not waive retention gates.

A bounded prospective diagnostic can replicate the four failed cells plus a
passing matched control with the same A/A + ABBA order and all per-sample
values. Its inputs, sample counts, controls and analysis must be declared before
capture. Such a diagnostic cannot retroactively accept 0726; a future retention
attempt still needs independently specified full-matrix qualification.
