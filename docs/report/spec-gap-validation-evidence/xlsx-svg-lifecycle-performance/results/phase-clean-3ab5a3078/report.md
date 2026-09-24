# XLSX SVG same-drawing phase decomposition

This is an exploratory descriptive view of one public attach operation
split into explicit phases. It reports absolute per-process medians and
ranges. It does not compare against the sealed baseline and makes no
speedup, regression, causal, or scaling claim.

Each phase resets the process-local counter baseline, while objects needed
by later phases remain live. `requested_alloc_bytes` is phase-local;
`live_after_bytes` and `retained_live_bytes_after` expose the retained
boundary. Phase values must not be summed or subtracted across unlike
live sets.

The six phase clocks are `open` (Workbook::from_bytes), `stages`
(Workbook::edit plus all public attach calls), `commit`, `firstsave`,
`reopen_secondsave`, and `validation` (the complete attach semantic
predicate, including graph, picture, opaque, and reopen-byte checks).

## `multi_picture_same_drawing_16`

| phase | elapsed median p1/p2/p3 (ns) | elapsed process range (ns) | requested allocation median p1/p2/p3 (bytes) | live delta median p1/p2/p3 (bytes) | retained live after median p1/p2/p3 (bytes) | RSS p1/p2/p3 (KiB) |
|---|---:|---:|---:|---:|---:|---:|
| `open` | 82515.0 / 81756.0 / 82190.5 | [80810,123010] | 557573.0 / 557573.0 / 557573.0 | 28046.0 / 28046.0 / 28046.0 | 46024.0 / 46024.0 / 46024.0 | 7496 / 7824 / 7696 |
| `stages` | 156445.5 / 153550.0 / 153645.5 | [151801,186561] | 104430.0 / 104430.0 / 104430.0 | 11373.0 / 11373.0 / 11373.0 | 57397.0 / 57397.0 / 57397.0 | 7496 / 7824 / 7696 |
| `commit` | 1887329.0 / 1874013.5 / 1878233.5 | [1857548,2019889] | 1172174.0 / 1172174.0 / 1172174.0 | 40535.0 / 40535.0 / 40535.0 | 97932.0 / 97932.0 / 97932.0 | 7496 / 7824 / 7696 |
| `firstsave` | 288991.5 / 282976.5 / 284586.0 | [278501,332912] | 633702.0 / 633702.0 / 633702.0 | 13216.0 / 13216.0 / 13216.0 | 111148.0 / 111148.0 / 111148.0 | 7496 / 7824 / 7696 |
| `reopen_secondsave` | 243391.0 / 241446.5 / 241796.0 | [239081,256312] | 728256.0 / 728256.0 / 728256.0 | 87925.0 / 87925.0 / 87925.0 | 147165.0 / 147165.0 / 147165.0 | 7496 / 7824 / 7696 |
| `validation` | 10613952.5 / 10549322.0 / 10609827.5 | [10521177,11041190] | 27730006.0 / 27730006.0 / 27730006.0 | 0.0 / 0.0 / 0.0 | 147165.0 / 147165.0 / 147165.0 | 7496 / 7824 / 7696 |

Largest observed phase medians by fresh process (descriptive only): p1: `validation`=10613952.5 ns; p2: `validation`=10549322.0 ns; p3: `validation`=10609827.5 ns.
Largest phase-local requested-allocation medians by fresh process (descriptive only): p1: `validation`=27730006.0 bytes; p2: `validation`=27730006.0 bytes; p3: `validation`=27730006.0 bytes.
RSS process range: [7496,7824] KiB; this is whole-process `/usr/bin/time -v` RSS.

## `multi_picture_same_drawing_64`

| phase | elapsed median p1/p2/p3 (ns) | elapsed process range (ns) | requested allocation median p1/p2/p3 (bytes) | live delta median p1/p2/p3 (bytes) | retained live after median p1/p2/p3 (bytes) | RSS p1/p2/p3 (KiB) |
|---|---:|---:|---:|---:|---:|---:|
| `open` | 102890.0 / 102530.5 / 102295.0 | [99350,134881] | 580307.0 / 580307.0 / 580307.0 | 50780.0 / 50780.0 / 50780.0 | 69510.0 / 69510.0 / 69510.0 | 8676 / 8576 / 8592 |
| `stages` | 568572.5 / 561367.0 / 565673.0 | [552383,582653] | 403438.0 / 403438.0 / 403438.0 | 44541.0 / 44541.0 / 44541.0 | 114051.0 / 114051.0 / 114051.0 | 8676 / 8576 / 8592 |
| `commit` | 6771140.0 / 6759905.5 / 6787200.0 | [6740110,7068781] | 4399110.0 / 4399110.0 / 4399110.0 | 136535.0 / 136535.0 / 136535.0 | 250586.0 / 250586.0 / 250586.0 | 8676 / 8576 / 8592 |
| `firstsave` | 837109.0 / 838579.0 / 836313.5 | [825793,868484] | 1087613.0 / 1087613.0 / 1087613.0 | 42944.0 / 42944.0 / 42944.0 | 293530.0 / 293530.0 / 293530.0 | 8676 / 8576 / 8592 |
| `reopen_secondsave` | 620068.0 / 621288.0 / 623888.0 | [611753,631223] | 1339070.0 / 1339070.0 / 1339070.0 | 285539.0 / 285539.0 / 285539.0 | 397993.0 / 397993.0 / 397993.0 | 8676 / 8576 / 8592 |
| `validation` | 131239422.0 / 131082651.5 / 131291267.5 | [130673905,131739239] | 205991190.0 / 205991190.0 / 205991190.0 | 0.0 / 0.0 / 0.0 | 397993.0 / 397993.0 / 397993.0 | 8676 / 8576 / 8592 |

Largest observed phase medians by fresh process (descriptive only): p1: `validation`=131239422.0 ns; p2: `validation`=131082651.5 ns; p3: `validation`=131291267.5 ns.
Largest phase-local requested-allocation medians by fresh process (descriptive only): p1: `validation`=205991190.0 bytes; p2: `validation`=205991190.0 bytes; p3: `validation`=205991190.0 bytes.
RSS process range: [8576,8676] KiB; this is whole-process `/usr/bin/time -v` RSS.

## `multi_picture_same_drawing_256`

| phase | elapsed median p1/p2/p3 (ns) | elapsed process range (ns) | requested allocation median p1/p2/p3 (bytes) | live delta median p1/p2/p3 (bytes) | retained live after median p1/p2/p3 (bytes) | RSS p1/p2/p3 (KiB) |
|---|---:|---:|---:|---:|---:|---:|
| `open` | 159861.0 / 156165.5 / 151855.0 | [133011,191281] | 671697.0 / 671697.0 / 671697.0 | 142170.0 / 142170.0 / 142170.0 | 163716.0 / 163716.0 / 163716.0 | 12628 / 12844 / 12296 |
| `stages` | 2183099.0 / 2173285.0 / 2167299.5 | [2132139,2298421] | 1599531.0 / 1599531.0 / 1599531.0 | 177213.0 / 177213.0 / 177213.0 | 340929.0 / 340929.0 / 340929.0 | 12628 / 12844 / 12296 |
| `commit` | 33889262.5 / 33784223.0 / 33762608.5 | [33609668,34693811] | 19640507.0 / 19640507.0 / 19640507.0 | 523357.0 / 523357.0 / 523357.0 | 864286.0 / 864286.0 / 864286.0 | 12628 / 12844 / 12296 |
| `firstsave` | 3094038.5 / 3134484.0 / 3092928.0 | [3061143,3223254] | 3031927.0 / 3031927.0 / 3031927.0 | 223744.0 / 223744.0 / 223744.0 | 1088030.0 / 1088030.0 / 1088030.0 | 12628 / 12844 / 12296 |
| `reopen_secondsave` | 2111145.0 / 2193504.5 / 2081964.0 | [2057719,2281010] | 3772642.0 / 3772642.0 / 3772642.0 | 1079015.0 / 1079015.0 / 1079015.0 | 1466475.0 / 1466475.0 / 1466475.0 | 12628 / 12844 / 12296 |
| `validation` | 1977560350.0 / 1989140384.5 / 1921114116.5 | [1909670050,2021523635] | 2435025383.0 / 2435025383.0 / 2435025383.0 | 0.0 / 0.0 / 0.0 | 1466475.0 / 1466475.0 / 1466475.0 | 12628 / 12844 / 12296 |

Largest observed phase medians by fresh process (descriptive only): p1: `validation`=1977560350.0 ns; p2: `validation`=1989140384.5 ns; p3: `validation`=1921114116.5 ns.
Largest phase-local requested-allocation medians by fresh process (descriptive only): p1: `validation`=2435025383.0 bytes; p2: `validation`=2435025383.0 bytes; p3: `validation`=2435025383.0 bytes.
RSS process range: [12296,12844] KiB; this is whole-process `/usr/bin/time -v` RSS.
