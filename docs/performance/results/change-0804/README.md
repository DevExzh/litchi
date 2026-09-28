# 0804 — bounded linear equality diagnostic

Before is the exact 0803 after-control. After changes only the bounded linear
name scan from byte ordering followed by equality testing to byte-slice
equality. Ordered-map comparison, empty-check placement, backend handoff,
key/error preflight, clone and fused behavior stay fixed. Test-only comparison
counting must still observe every name comparison. Neither leg is current
production; this diagnostic cannot authorize workflow advancement or adoption.

The hypothesis is that equality can avoid computing ordering in the bounded
scan, including when lengths differ. A gain is not assumed. The same 39 literal
cases and construct/consume owners run in one fresh release binary. Six paired
native blocks use 30 samples, three warmups and 4,096 iterations on CPU 12.
Nearest-rank process p50s and a 10,000-resample paired-median bootstrap use
fresh seed 804080. Two separate Callgrind repeats retain guest instruction
and branch diagnostics. Historical native timings are not pooled.

The one source intervention includes associated compiler/code-layout effects.
It does not isolate a particular machine instruction as the native-time cause.
Construction includes an opaque iterator reference, drop, dispatch and checksum
work; consumption includes construction and checksum work. Object size is
measured separately and is not heap use. No public-workflow, resource or
cross-format benefit follows from these hot micro-inputs.

Root owns serialized builds and captures. The source/architecture witnesses
separately preserve production. Archive-relative `candidate.patch` applies
only to `candidate/before`, never to production. Independent source review,
helper mirror tests, semantic/clone oracles and offline replay verify the
comparison. Actual failed attempts, if any, are retained; the owned target is
removed only after its executable identity is verified.

## Final diagnostic result

Sixteen/32/33-name consumption improves by 15.589%/24.825%/24.259% against the
0803 after-control. No row triggers the frozen diagnostic regression rule;
short-row increases, four construction medians above 5% and all 75 spread flags
remain visible. Both layouts measure 128 bytes. Neither leg is production,
and advancement/adoption remain false.

Both mirror legs pass 100 tests and Clippy. The initial root metadata omission
is retained in `execution-note.txt`. The original profile selector failed after
one zero-exit child produced only an empty termination dump. That attempt is
retained under `profiles-failed-0`; it is not counted as a successful capture.
`profile-owner-resolution.json` and a separate driver bind both construction
legs to the surviving `after_construct` symbol. The binary, native reports,
original plan and build inputs did not change.

The successful lanes contain 1,248 reports and 28,392 samples; the failed scope
adds one separate report/sample and a zero dump. All 312 resolved profile owners
qualify and all five counters conserve across 624 successful dumps. See the
main report, `summary.md`, `decision.json`, source/results reviews and the
failure audit for individual rows and limits.

Offline replay after cleanup:

```sh
python3 -B docs/performance/results/change-0804/validate.py --require-final-seal
python3 -B docs/performance/results/change-0804/seal_packet.py --check-head
```
