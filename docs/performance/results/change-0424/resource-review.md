# Resource result and retention decision

Keep `d18bf7db4` for the measured media lifecycle resource benefit. Both
30-sample allocator repeats give the same request counts/bytes and the same
mean region peaks. All 16 exact-corpus/output captures pass. The source and
ownership review and 86 applicable Rust tests support unchanged preservation
and refusal behavior. This is a scoped resource diagnostic; no registered
release latency claim is added.

| Operation metric | Plain control → candidate | Media control → candidate |
| --- | ---: | ---: |
| Requested bytes | 8,916,495 → 8,916,495 | 117,608,609 → 100,831,137 (−14.266%) |
| Allocation calls | 10,993 → 10,993 | 13,314 → 13,306 (−0.060%) |
| Mean region peak live bytes | 1,059,128 → 1,059,132 | 229,300,275 → 212,522,935 (−7.317%) |

The media region peak falls by 16,777,340 bytes; requested volume falls by
16,777,472 bytes. The separate Heaptrack lifecycle ancestry shows the mechanism:
the original eight-image, 16,777,216-byte planning clone remains, while the
publication clone is absent in the candidate. Whole-command Heaptrack totals
also include several validation/oracle lifecycles and cannot substitute for
the operation vectors above. Request volume is allocator callback accounting,
not a measurement of physical copied bytes or memory bandwidth.

Region peak is the absolute process live-byte maximum observed during the
serialized operation, including entry live bytes, under
`serialized_region_peak_v3`. It is distinct from the process lifetime peak,
whole-process RSS and memory owned by any one snapshot. Media live-after mean
increases by 196 bytes; entry-live and process lifetime peak each increase by
4 bytes. The large transient reduction therefore does not establish a reduction
in retained endpoint state. The plain 4-byte peak difference is retained rather
than rounded into an exact equality claim. All underlying vectors remain in
`matched/summary.json` and the original reports.

Normal timing has no regression above 5%. Plain p50 is +0.805% / +0.755%, mean
+0.863% / +0.797%, p95 +0.689% / +0.841%, and p99 +1.398% / +0.559%.
These small adverse observations are accepted for the repeatable media resource
benefit and remain visible individually. Media p50 is −0.013% / −3.825%, but
control repeats drift +4.025% in p50 while candidate repeats drift +0.059%.
All four groups meet the predeclared 5% p50/mean, 10% p95 and 15% p99 ceilings;
the differing repeat drift prevents a convincing causal media timing conclusion.
There is no speedup claim from these 100-sample diagnostic legs.

Whole-process RSS is effectively unchanged. Normal media pairs are −0.030% /
−0.191%; allocator media pairs are +0.015% / +0.033%. The largest adverse RSS
pair is plain allocator R1 at +0.107%. Setup includes a large generated corpus
and correctness work, so the roughly 803,000 KiB media process peak is not the
operation-region peak. Small endpoint `rss_delta_bytes` values include zeros;
the R1 percentage of −100% is not a process-memory reduction claim.

Logical read counts and bytes are unchanged between roles: plain source
123 reads / 14,105 bytes and destination 423 / 58,432; media source
1,794 / 50,368,324 and destination 975 / 16,845,847. These are counted
in-memory `ReadAt` calls with adapter overhead, not filesystem syscalls or
remote requests. Full source reads, decompression, exact byte comparisons,
graph validation and candidate reread remain. Identical output digests and
accepted sink sizes bind every run: 31,514 bytes plain and 33,599,843 media.

The new Arc handles refer to independently staged immutable payloads. Full
original and publication staging reservations and the candidate reread charge
remain conservative. The plan stays borrowed through publication. No managed
budget reduction, cache-retention, post-drop, native-producer performance,
concurrency/scaling or physical-I/O claim follows. The work is small and
measured; the remaining work is listed in `next-work.md`.
