# 0783 lifecycle phase summary

Process p50 values use nearest-rank p50 over 30 measured samples; phase shares are medians within each process and then medians across six alternating processes.

| Shape | Leg | Total p50 (ns) | Capture | Stage | Commit | Apply | Serialize | RSS p50 (KiB) | Total spread >5% | RSS spread >5% |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- | --- |
| tiny | control | 1434526.500 | — | — | — | — | — | 4970.000 | False | True |
| tiny | phases | 1422951.000 | 17.233% | 4.399% | 15.233% | 2.806% | 60.115% | 5132.000 | False | True |
| medium | control | 2058744.000 | — | — | — | — | — | 5248.000 | False | False |
| medium | phases | 2040659.000 | 23.908% | 4.840% | 14.778% | 2.749% | 53.454% | 5444.000 | False | True |
| large | control | 31986500.000 | — | — | — | — | — | 16552.000 | False | False |
| large | phases | 30720849.000 | 69.051% | 2.671% | 4.290% | 0.631% | 23.320% | 16592.000 | False | False |

| Shape | Phases/control total p50 ratio | Change | 95% bootstrap CI | Ratio spread >5% |
| --- | ---: | ---: | --- | --- |
| tiny | 0.990512 | -0.949% | [0.989173, 0.996081] | False |
| medium | 0.990573 | -0.943% | [0.989369, 0.992642] | False |
| large | 0.963472 | -3.653% | [0.958465, 0.964104] | False |
