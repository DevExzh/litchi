# 0456: retain shared ZIP payloads during publication

The preservation writer previously copied each verified compressed payload into
a temporary one-member ZIP archive, then parsed that archive to retain its local
and central ranges. The new path retains the verified Store/Deflate token or
shared Store Arc and prepares only local framing and central metadata. The final
forward output pass writes the retained payload directly. This removes a
payload-sized preparation allocation and copy without changing the output.
Ordinary owned Store and generated Deflate payloads keep their buffered path.

Both writer paths share checked local/central framing grammar. Full output
layout preflight, name validation before shared Store CRC work, ZIP64 offset
promotion and accepted-byte accounting remain in place. The managed OPC budget
still includes its conservative writer-payload allowance. This batch measures
physical allocation changes; it does not lower admission budgets or claim a
bounded complete document lifecycle. The 64 KiB stack copy buffer is unchanged.

The frozen matched experiment contains 24 ABBA lanes, 30 samples after 3 warmups,
CPU 2 and one worker: 480 ordinary observations and 240 separate allocator
observations. The range provider simulates a 64 KiB request cap, 200 us per
request and 25 MiB/s with separate sleeps. It does not measure physical network,
storage, or cold-cache behavior. All absolute >5% timing/RSS comparisons remain
visible, including improvements and adverse observations.

Media publication allocation is identical in both repeats: calls fall from
5,397 to 5,381; allocated bytes from 21,057,867 to 4,243,083 (-79.85%); and
regional heap peak above operation entry from 17,406,388 to 593,892 (-96.59%).
Net live growth is unchanged. Plain publication calls remain 4,251; allocated
bytes rise 3,360 (+0.082%) and peak above entry rises 1,720 (+0.309%), consistent
with the larger prepared-entry metadata representation. Process RSS changes
stay within 4.611%; the whole harness peak is dominated by other retained work.

Output bytes, hashes, provider work, sink writes and managed budgets remain
identical. Media publication still reads 16,830,603 destination bytes in 577
calls and emits 33,599,715 bytes in 767 sink writes. The publication source-read
count is zero. This is a preparation allocation/copy saving, not an I/O saving.

The ordinary bytes/media comparison is adverse in both repeats: API medians
rise from 25.071/25.050 ms to 28.935/29.145 ms (+15.411%/+16.347%); publication
rises from 12.523/12.532 ms to 16.607/16.638 ms (+32.611%/+32.769%). Candidate
minor faults are 922,403/835,710 versus baseline 383,775/385,304. All 15 absolute
>5% flags remain retained: 12 positive bytes/media API/publication percentiles,
one range/media destination-open p99 increase (+10.210%, not repeated), and two
range/media open-tail improvements. Range API medians change -0.051%/-0.385%;
plain API medians stay within 0.512%. No unconditional latency gain is claimed.

Release checks pass 460 ZIP, 462 ODF-common, 497 OPC, 854 PPTX and 381 performance
harness tests: 2,654 total, 7 ignored. Strict lint, documentation, formatting,
minimal non-iWork workspace and crate-boundary gates pass. The extended ZIP
ASAN fuzz target passes 1,000 runs using six retained deterministic Store/Deflate
seeds. Focused tests verify storage ownership, exact framing, partial payload and
central-directory failures, and synthetic ZIP64 offset promotion. The initial
writer reborrow compile failure is retained with its exact source and successful
subsequent gate; no failed measurement was discarded.

The pinned LibreOffice QA self-pair emits the exact previous 42,948-byte golden
output through bytes and range providers. ZIP CRC, member order, copied
slide/image bytes, metadata hashes and XML parsing pass. The first metadata
check failed because a broad batch-number substitution altered a SHA string in
the copied expected file. The corrected attempt uses a byte-identical copy of
the prior expected file and fresh output paths; the original failure is retained.
No golden output or production source was changed to obtain a pass.

Separate whole-process perf stat runs have about 381 thousand faults for both
builds. Mean candidate instructions change -0.078%, cycles -1.015% and cache
misses -5.192%; these include corpus generation, output oracles and serialization,
and do not attribute hardware counters to the timed API. Local instruction
profiles contain 972/971 samples (about 48.60/48.55 billion instructions), with
corpus compression and SHA-256 dominant. Kernel symbol restrictions and partial
unwinding limit attribution. No API-level CPU claim follows from these profiles.

A separately frozen eight-process diagnostic fixes glibc's mmap threshold to
128 KiB and 32 MiB, with ABBA at each setting. Explicit thresholds disable its
dynamic adjustment, as described in the
[glibc manual](https://sourceware.org/glibc/manual/latest/html_node/Malloc-Tunable-Parameters.html).
At 128 KiB, candidate API medians improve 9.999%/10.266% and process faults fall
from about 1.779 million to 1.631 million. At 32 MiB, API medians improve
2.626%/2.989%; one baseline process still has more faults. These 240 diagnostic
samples remain separate from the 720 formal observations. They show allocator
policy materially affects the comparison, without reconstructing the original
mapping decisions or erasing the original ordinary regression.

Retain the change for the repeatable 79.85% allocation and 96.59% regional-peak
reduction, with the observed allocator-sensitive latency tradeoff disclosed.
The code does not tune the process allocator. Further deployment-representative
measurement is needed before claiming a general bytes-provider speedup.

Six distinct original native PPTX pair probes all refuse incompatible
layout/master/theme graphs. Those actual refusals remain in the evidence and
are not successful copy or native-application results. Broader fixture matching
still needs work. The next bounded existing-document append slice needs a
streaming common XML transform and format-owned semantics; the current ODP
source presentation and replacement publisher retain complete XML. This ZIP
optimization does not remove those allocations. No coverage taxonomy row is
promoted, and the full non-iWork goal remains open.

Actual precleanup verification and separate-copy portable replay pass. The
owned task directory is removed after inventorying 373 files / 522,511,229 bytes;
shared Cargo target caches remain. The sealed
[evidence bundle](../results/change-0456/README.md) retains source snapshots,
receipts, failed attempts, raw reports, derived results and replay drivers.
