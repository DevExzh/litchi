# Matched median statistics

This file compares the accepted corrected candidate against the pinned baseline using 15 fresh-process samples and 3 warmups for each of 15 groups (450 raw samples). Ratios are candidate median divided by baseline median; values above 1 indicate a higher observed candidate value. These are descriptive observations from a shared host. They do not establish tail behavior, statistical certainty, or asymptotic complexity.

The candidate source has the same detached Git base commit string as the baseline because the selected source is a frozen working tree. The freeze receipt has SHA-256 `f19c50c75401cc9193d22d5605a0376720630e4185d15717db93338cdb8f39ec` and hashes 11 selected files; the corrected `embedded_data.rs` hash is `0d0a8407724d4ba4c53e2613b547ff227268ac88f8b3e3003bd5d6038663a815`. `gates/freeze.json` and its selected-file hashes identify the candidate contents. Both captures use source Cargo.lock `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, harness Cargo.lock `01146937be7eac11b4484aeb2fcf6ee181c157c47973e21a67b7c15fdd576b29`, identical profile manifests, compiler, flags, fixture bytes, and process settings.

## elapsed (ms)

| size | lane | baseline median | candidate median | ratio | delta |
| --- | --- | ---: | ---: | ---: | ---: |
| small | read | 0.803644 | 0.946294 | 1.1775 | +17.8% |
| small | noop-commit-save | 1.587338 | 1.874158 | 1.1807 | +18.1% |
| small | payload-replacement | 3.078105 | 3.851728 | 1.2513 | +25.1% |
| small | rename-binding-rewrite | 3.174015 | 3.950798 | 1.2447 | +24.5% |
| small | remove-inverse | 4.692472 | 5.837457 | 1.2440 | +24.4% |
| medium | read | 5.626517 | 6.444250 | 1.1453 | +14.5% |
| medium | noop-commit-save | 11.201294 | 12.884601 | 1.1503 | +15.0% |
| medium | payload-replacement | 22.138217 | 26.200373 | 1.1835 | +18.3% |
| medium | rename-binding-rewrite | 22.648829 | 26.707595 | 1.1792 | +17.9% |
| medium | remove-inverse | 34.325415 | 40.275049 | 1.1733 | +17.3% |
| large | read | 22.926921 | 26.038282 | 1.1357 | +13.6% |
| large | noop-commit-save | 46.480054 | 52.818458 | 1.1364 | +13.6% |
| large | payload-replacement | 91.217660 | 106.112932 | 1.1633 | +16.3% |
| large | rename-binding-rewrite | 92.902609 | 107.863121 | 1.1610 | +16.1% |
| large | remove-inverse | 141.946988 | 164.784839 | 1.1609 | +16.1% |

## requested allocation bytes

| size | lane | baseline median | candidate median | ratio | delta |
| --- | --- | ---: | ---: | ---: | ---: |
| small | read | 949575 | 1021962 | 1.0762 | +7.6% |
| small | noop-commit-save | 1918354 | 2063128 | 1.0755 | +7.5% |
| small | payload-replacement | 4196572 | 4588950 | 1.0935 | +9.3% |
| small | rename-binding-rewrite | 4250631 | 4653641 | 1.0948 | +9.5% |
| small | remove-inverse | 5510520 | 6094591 | 1.1060 | +10.6% |
| medium | read | 7153519 | 7544109 | 1.0546 | +5.5% |
| medium | noop-commit-save | 14366009 | 15147189 | 1.0544 | +5.4% |
| medium | payload-replacement | 28438013 | 30461190 | 1.0711 | +7.1% |
| medium | rename-binding-rewrite | 28866398 | 30897295 | 1.0704 | +7.0% |
| medium | remove-inverse | 42078167 | 45197703 | 1.0741 | +7.4% |
| large | read | 30439339 | 31934553 | 1.0491 | +4.9% |
| large | noop-commit-save | 61078721 | 64069149 | 1.0490 | +4.9% |
| large | payload-replacement | 118819322 | 126522899 | 1.0648 | +6.5% |
| large | rename-binding-rewrite | 118475216 | 126176529 | 1.0650 | +6.5% |
| large | remove-inverse | 175944228 | 187895356 | 1.0679 | +6.8% |

## allocation calls

| size | lane | baseline median | candidate median | ratio | delta |
| --- | --- | ---: | ---: | ---: | ---: |
| small | read | 9087 | 10082 | 1.1095 | +10.9% |
| small | noop-commit-save | 18408 | 20398 | 1.1081 | +10.8% |
| small | payload-replacement | 35770 | 41177 | 1.1512 | +15.1% |
| small | rename-binding-rewrite | 36193 | 41710 | 1.1524 | +15.2% |
| small | remove-inverse | 53826 | 61871 | 1.1495 | +14.9% |
| medium | read | 67875 | 73107 | 1.0771 | +7.7% |
| medium | noop-commit-save | 136530 | 146994 | 1.0766 | +7.7% |
| medium | payload-replacement | 271003 | 297626 | 1.0982 | +9.8% |
| medium | rename-binding-rewrite | 273428 | 300161 | 1.0978 | +9.8% |
| medium | remove-inverse | 412490 | 453386 | 1.0991 | +9.9% |
| large | read | 278284 | 298018 | 1.0709 | +7.1% |
| large | noop-commit-save | 559220 | 598688 | 1.0706 | +7.1% |
| large | payload-replacement | 1112991 | 1212222 | 1.0892 | +8.9% |
| large | rename-binding-rewrite | 1122280 | 1221621 | 1.0885 | +8.9% |
| large | remove-inverse | 1704192 | 1857546 | 1.0900 | +9.0% |

## peak live bytes

| size | lane | baseline median | candidate median | ratio | delta |
| --- | --- | ---: | ---: | ---: | ---: |
| small | read | 297902 | 298318 | 1.0014 | +0.1% |
| small | noop-commit-save | 316806 | 317612 | 1.0025 | +0.3% |
| small | payload-replacement | 513617 | 514397 | 1.0015 | +0.2% |
| small | rename-binding-rewrite | 525472 | 526252 | 1.0015 | +0.1% |
| small | remove-inverse | 332849 | 334045 | 1.0036 | +0.4% |
| medium | read | 2292248 | 2292664 | 1.0002 | +0.0% |
| medium | noop-commit-save | 2459575 | 2460381 | 1.0003 | +0.0% |
| medium | payload-replacement | 2508860 | 2510056 | 1.0005 | +0.0% |
| medium | rename-binding-rewrite | 2591391 | 2592587 | 1.0005 | +0.0% |
| medium | remove-inverse | 2548532 | 2549728 | 1.0005 | +0.0% |
| large | read | 10676156 | 10676572 | 1.0000 | +0.0% |
| large | noop-commit-save | 12121003 | 12121809 | 1.0001 | +0.0% |
| large | payload-replacement | 13115084 | 13116280 | 1.0001 | +0.0% |
| large | rename-binding-rewrite | 12477471 | 12478667 | 1.0001 | +0.0% |
| large | remove-inverse | 13275572 | 13276768 | 1.0001 | +0.0% |

## maximum RSS (KiB)

| size | lane | baseline median | candidate median | ratio | delta |
| --- | --- | ---: | ---: | ---: | ---: |
| small | read | 4316 | 4556 | 1.0556 | +5.6% |
| small | noop-commit-save | 4504 | 4532 | 1.0062 | +0.6% |
| small | payload-replacement | 4800 | 5060 | 1.0542 | +5.4% |
| small | rename-binding-rewrite | 5032 | 5076 | 1.0087 | +0.9% |
| small | remove-inverse | 4552 | 4572 | 1.0044 | +0.4% |
| medium | read | 6860 | 6844 | 0.9977 | -0.2% |
| medium | noop-commit-save | 7120 | 7112 | 0.9989 | -0.1% |
| medium | payload-replacement | 7356 | 7364 | 1.0011 | +0.1% |
| medium | rename-binding-rewrite | 7632 | 7620 | 0.9984 | -0.2% |
| medium | remove-inverse | 7372 | 7356 | 0.9978 | -0.2% |
| large | read | 17840 | 17824 | 0.9991 | -0.1% |
| large | noop-commit-save | 19640 | 20408 | 1.0391 | +3.9% |
| large | payload-replacement | 21516 | 21748 | 1.0108 | +1.1% |
| large | rename-binding-rewrite | 20676 | 20908 | 1.0112 | +1.1% |
| large | remove-inverse | 20656 | 20928 | 1.0132 | +1.3% |

The corrected candidate removes the discarded reserve-ceiling behavior by growing the connection-edge vector one edge at a time with fallible `try_reserve(1)`. The prior pre-reserve-fix capture is retained under `candidate-pre-reserve-fix` with a reconstruction patch and is excluded from every accepted baseline/candidate comparison.

The source package, Custom Data operation, XLSX serialization, and result reopening are inside the timed boundary. Allocator metrics include the observer overhead; peak live bytes are aggregate allocator accounting, while RSS is from `/usr/bin/time -v`.
