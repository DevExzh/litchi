# change-0492 formal1 results review

This is an independent read-only review of the 16 retained `formal1` child
receipts. The frozen protocol and final2 build receipts loaded successfully;
all 16 terminals report `pass`, exit code 0, no timeout or termination, source
unchanged, and private cleanup `pass`. Each capture has the exact seven-file
inventory (`started.json`, `terminal.json`, `report.json`, `resource.txt`,
`stdout.txt`, `stderr.txt`, and `replay-cleanup.json`). The raw reports validate
as `docx_provider_lifecycle_v2` with 30 measured rows after warmup 3, for 480
measured samples.

The custody bindings are source revision
`e44a23396146d504ffc738e0989de896635f02a3`, source manifest
`18fbdb0df7adc7a2482fbd9e458bcabd7d79581f8842c8e01038e6aeabb07a2b`, protocol
`9c0d7364d44eb97ca6905b1b9b4592acdb1001e58005098dc237a414691af0e2`, normal
binary `8659d4f40f6e11376da186818187488a5ea43c37f19a1cc996307e8618ed5440`,
and allocator binary
`e7820bc6aae1f6d43b2b17ddbcaac656de6ec08440b53d5e0ca412273bea8a3a`.

All latency values below are nanoseconds. Each pair compares the exact-read
baseline under a 64 KiB transport cap with the 4 KiB candidate within the same
role, repeat, and delay arm.
The bootstrap interval is the deterministic 2,000-resample percentile 95%
interval for the median. Percent changes are `(candidate - baseline) /
baseline`, and a flag means a positive latency change above 5% at that
percentile.

| role | repeat | service arm | baseline p50 / p95 / p99 | candidate p50 / p95 / p99 | baseline median bootstrap | candidate median bootstrap | candidate change p50 / p95 / p99 | latency flag |
| --- | ---: | --- | --- | --- | --- | --- | --- | --- |
| normal | 1 | 0 us | 277,391 / 285,941 / 288,521 | 290,917 / 299,901 / 303,532 | 276,851–280,101 | 289,346–294,226 | +4.876% / +4.882% / +5.203% | p99 |
| normal | 1 | 1,000 us + 100 MiB/s | 20,343,529 / 20,549,819 / 20,688,090 | 3,465,600 / 3,477,125 / 3,507,266 | 20,323,333–20,356,644 | 3,464,495–3,469,335 | −82.965% / −83.080% / −83.047% | none |
| normal | 2 | 0 us | 274,711 / 286,152 / 296,681 | 276,081 / 286,162 / 291,601 | 274,101–278,386 | 275,582–277,716 | +0.499% / +0.003% / −1.712% | none |
| normal | 2 | 1,000 us + 100 MiB/s | 20,335,863 / 20,406,659 / 20,475,219 | 3,451,745 / 3,487,796 / 4,603,770 | 20,323,179–20,343,268 | 3,449,470–3,457,370 | −83.026% / −82.909% / −77.515% | none |
| allocator | 1 | 0 us | 291,511 / 299,342 / 300,982 | 290,176 / 301,111 / 302,022 | 290,241–293,991 | 288,646–293,136 | −0.458% / +0.591% / +0.346% | none |
| allocator | 1 | 1,000 us + 100 MiB/s | 20,340,699 / 20,423,249 / 20,515,959 | 3,459,380 / 3,469,255 / 3,470,535 | 20,329,438–20,354,113 | 3,458,185–3,462,220 | −82.993% / −83.013% / −83.084% | none |
| allocator | 2 | 0 us | 291,506 / 299,151 / 301,112 | 291,546 / 302,412 / 304,272 | 290,186–294,296 | 290,721–295,241 | +0.014% / +1.090% / +1.049% | none |
| allocator | 2 | 1,000 us + 100 MiB/s | 20,338,373 / 20,381,239 / 20,567,220 | 3,461,065 / 3,469,335 / 3,509,235 | 20,329,343–20,348,758 | 3,459,560–3,465,550 | −82.983% / −82.978% / −82.938% | none |

The read-ahead contract is identical in all eight candidate process rows:
window capacity 4,096 bytes, 19 logical requests, 16 hits, 3 misses, 3
fills, 5,445 requested and returned fill bytes, maximum fill 4,096 bytes,
zero short fills, and zero failures. The baseline has the pinned 19 physical
calls and 3,966 requested bytes. The candidate's physical trace has 3 calls
and 5,445 requested/returned bytes. The candidate physical trace overlaps 69
compressed media bytes from the metadata fill; the logical trace overlaps zero,
so this evidence makes no zero-physical-overlap claim. The 4 KiB window and
its fresh buffer setup are outside the operation clock as specified by the
protocol.

The allocator receipts were read directly because the frozen analyzer's raw
vector names are `allocation_allocation_calls` and
`allocation_reallocation_calls`, while its generic comparison lookup uses the
shorter names. The standalone mapping in
`results-review.json` preserves the actual per-row values and report hashes.
The following table shows every allocator process; all p50/p95/p99 values are
reported as absolute region peak or peak increment bytes, and the arrows are
baseline → candidate within the matching delay arm.

| repeat | service arm | allocation calls | reallocation calls | allocated bytes | region peak p50 / p95 / p99 | peak increment p50 / p95 / p99 |
| ---: | --- | --- | --- | --- | --- | --- |
| 1 | 0 us | 901 → 902 | 116 → 116 | 653,366 → 653,462 | 19,056,676 / 19,069,852 / 19,070,828 → 19,055,860 / 19,063,852 / 19,064,444 | 139,203 / 139,203 / 139,203 → 139,299 / 139,299 / 139,299 |
| 1 | 1,000 us + 100 MiB/s | 901 → 902 | 116 → 116 | 653,366 → 653,462 | 19,056,913 / 19,070,089 / 19,071,065 → 19,056,097 / 19,064,089 / 19,064,681 | 139,203 / 139,203 / 139,203 → 139,299 / 139,299 / 139,299 |
| 2 | 0 us | 901 → 902 | 116 → 116 | 653,366 → 653,462 | 19,056,676 / 19,069,852 / 19,070,828 → 19,055,860 / 19,063,852 / 19,064,444 | 139,203 / 139,203 / 139,203 → 139,299 / 139,299 / 139,299 |
| 2 | 1,000 us + 100 MiB/s | 901 → 902 | 116 → 116 | 653,366 → 653,462 | 19,056,913 / 19,070,089 / 19,071,065 → 19,056,097 / 19,064,089 / 19,064,681 | 139,203 / 139,203 / 139,203 → 139,299 / 139,299 / 139,299 |

The absolute region peak includes the existing live allocation baseline. The
peak increment isolates the operation's added region usage: the candidate is
96 bytes higher in this run, with one additional allocation call and no change
in reallocation calls. None of these allocation changes crosses the 5%
adverse threshold.

`/usr/bin/time -v` captured one whole-child maximum RSS observation per
process, including setup. The canonical analyzer represents that singleton
with identical p50/p95/p99 values; those quantiles are not an empirical
within-process RSS distribution.

| role | repeat | service arm | baseline RSS KiB | candidate RSS KiB |
| --- | ---: | --- | --- | --- |
| normal | 1 | 0 us | 93,988 | 93,360 |
| normal | 1 | 1,000 us + 100 MiB/s | 94,232 | 94,224 |
| normal | 2 | 0 us | 93,960 | 93,980 |
| normal | 2 | 1,000 us + 100 MiB/s | 94,152 | 93,332 |
| allocator | 1 | 0 us | 93,664 | 93,756 |
| allocator | 1 | 1,000 us + 100 MiB/s | 94,388 | 93,984 |
| allocator | 2 | 0 us | 93,748 | 93,956 |
| allocator | 2 | 1,000 us + 100 MiB/s | 93,716 | 93,584 |

The complete adverse list is structural and is retained in the canonical
analysis. Every role and repeat flags candidate physical requested and
returned bytes at p50, p95, and p99: 5,445 versus 3,966 is +37.292%. This is
the measured cost of the bounded physical overfetch and agrees with the 69-byte
candidate media overlap. The only latency adverse flag is normal repeat 1,
zero-delay p99 at +5.203%; its p50 and p95 changes are below 5%. There are no
allocator latency, RSS, or allocation adverse flags. The delayed latency rows
are much lower for the candidate because the frozen 1,000 us and transfer
pacing are charged for three candidate adapter calls versus 19 baseline calls;
they are descriptive evidence of this configured transport and do not support
a cross-arm optimization claim.

The timeout unit test verifies that `_kill_group` terminates a new process
group and returns after the child exits. It does not synthesize a retained
failed capture receipt; the capture path itself retains failed terminal state
when launch, timeout, validation, or cleanup fails. The result remains
descriptive (`claim_authorized: false`) and the synthetic range transport is
not a disk or network I/O measurement.
