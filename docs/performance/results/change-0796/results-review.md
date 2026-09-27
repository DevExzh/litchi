# 0796 results review

Status: passed for the stated diagnostic scope. I found no numerical, custody, or scope blocker in the final report and retained evidence.

## Evidence reviewed

I reviewed the final report, `summary.md`, `analysis.json`, the frozen fixture/catalog checks, the native and Callgrind raw evidence, the independent audit outputs, and the cleanup witness. The review was source/read-only and did not rebuild or rerun native captures.

The packet contains 744 native child reports with 22,320 samples and 248 profile children, represented by 248 positive Callgrind dumps and 248 termination dumps. The frozen catalog and source/lock checks pass. All native children completed successfully with no retry or discard path.

## Independent numerical agreement

The independent native audit verified every receipt, input checksum, semantic marker, accepted sequence, error marker, and checksum against the expected fixture data. It reconstructed the six paired p50 ratios and the seeded 10,000-replicate bootstrap intervals from the raw samples. The 39 diagnostic flags are exactly the 31 construction rows plus 8 consumption rows reported by the analyzer. The independently recomputed values for all 62 case/mode rows match the analyzer and the report, including the representative distinct, duplicate-long, and syntax-tail rows.

The report states the native-spread limitation accurately: 66 of 124 case/mode/leg groups exceed 5% p50 spread, no samples or blocks were dropped, and the bootstrap covers the six observed paired blocks rather than removing host or code-generation uncertainty. This supports diagnostic interpretation only.

## Callgrind and claim scope

The independent raw parser checked all 496 dumps. For every dump, the five reported events (`Ir`, `Bc`, `Bcm`, `Bi`, and `Bim`) conserve from the self summary through the file summary and packet totals; positive dumps have the expected owner section and termination dumps are zero with the termination marker. All 248 positive profiles qualify the named exact owner with one incoming call and a conserved self-plus-child partition. The analysis correctly retains inclusive nested costs and makes no allocator-API or fraction-of-total claim.

The report’s interpretation is bounded to this local, release-mode diagnostic binary and its hot microprobe. It preserves the opaque construction difference, same-binary/local-module caveat, timer/function-pointer/checksum overhead, hot-input limitation, guest-counter interpretation, and the absence of public workflow, allocation, latency, or adoption claims. The all-tail error layout and lack of a recovery-tail fixture are stated limitations; the earlier 0794 rejection remains authoritative.

## Custody and conclusion

The cleanup witness records removal of the exact captured binary and 112,130,069 target bytes, with the target absent afterward. The post-cleanup validation accepts the retained binary identity witness, and unrelated workspace files/worktrees remain outside the owned target. Captures, frozen inputs, lockfile, analysis, audits, and review evidence remain available in the packet.

The final report is numerically consistent with the raw evidence and independently reproduced audits, and its conclusions stay within the declared diagnostic scope. There is no remaining results-review blocker to sealing the packet.
