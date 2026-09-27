# 0798 diagnostic summary

The offline replay passes 45 reports and 45 samples: 15 plain controls and 30
instrumented census reports. Both census repeats are byte-for-byte equal in
their grouped rows, counters, and retained source bytes. Plain and instrumented
reports match the sealed 0794 semantic and publication identities after
excluding elapsed time and probe metadata.

The census covers `litchi-opc::xml_attributes::CheckedAttributes` instances on
the probe's calling thread during the generated PPTX capture, commit, and
lifecycle regions. It does not measure native latency, allocations, RSS,
instructions, or production performance.

| observation across both repeats | count |
| --- | ---: |
| iterator instances | 295,600 |
| `next` calls | 734,400 |
| successful `Ok` yields | 438,800 |
| error yields | 0 |
| `None` yields | 295,600 |
| full consumptions | 295,600 |
| early drops | 0 |
| never-advanced drops | 0 |
| clones | 0 |
| live instances at finish | 0 |

The exact lexical attribute-count histogram across both repeats is:

```text
0: 12,520; 1: 136,896; 2: 143,208; 3: 936; 4: 320;
6: 1,080; 7: 480; 9: 120; 12: 40
```

The per-case totals are identical in the two repeats:

| shape | operation | instances | attributes |
| --- | --- | ---: | ---: |
| tiny | capture | 429 | 750 |
| tiny | commit | 539 | 750 |
| tiny | lifecycle | 1,016 | 1,596 |
| medium | capture | 1,023 | 1,677 |
| medium | commit | 770 | 975 |
| medium | lifecycle | 1,889 | 2,844 |
| large | capture | 61,327 | 92,485 |
| large | commit | 5,250 | 4,603 |
| large | lifecycle | 67,777 | 99,488 |
| vendor | capture | 1,119 | 2,325 |
| vendor | commit | 778 | 1,137 |
| vendor | lifecycle | 1,993 | 3,654 |
| unicode-vendor | capture | 1,119 | 2,325 |
| unicode-vendor | commit | 778 | 1,137 |
| unicode-vendor | lifecycle | 1,993 | 3,654 |

Instance conservation, zero live instances, unsaturated counters, lineage
identity, aggregation conservation, independent lexical scans, and repeat
equality all pass. The full tag/source histogram and per-case grouped rows are
retained in [analysis.json](analysis.json); the independent lexical and
semantic cross-check is in [root-audit.json](root-audit.json).

This evidence supports examining a short-tag candidate that avoids replaying
the first attribute. It does not authorize production adoption or predict a
workflow speedup.
