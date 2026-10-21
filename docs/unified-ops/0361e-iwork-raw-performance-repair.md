# 0361e — Repair the IWA raw-operation performance regressions

- Date: 2026-10-21.
- Parent wave: [0361d](0361d-iwork-raw-limits.md).
- Baseline: `e63a1fe3534bcaaae2ce367339f7f8275877baef`.
- Status: repair and targeted native comparison accepted; final integrated gate pending.

The first 0361d comparison is rejected. Canonical framing takes 1.073× base,
and small exact catalog copy/write takes 1.599×/1.588×. The faster output-bound
calculation and unchanged bytes/allocations do not excuse those regressions.
All initial paired blocks and the discarded candidate's provenance remain in
the [first receipt](results/0361d/native-v1.json).

Catalog output now admits the actual artifact size directly. Its private
profile fields and validated constructors already establish profile validity;
the output limit still fails before allocation or sink access. Snappy keeps
its complete framing/resource preflight, exact output reservation, per-frame
initialization and raw payload validation. The second traversal reuses the
immutable input ranges admitted by the first traversal. A regression test
requires malformed raw payloads to fail even when their framing and declared
decoded lengths pass preflight.

Canonical framing detects an overlong unsigned varint from its final zero
group. The admitted per-object count and `u32` message lengths bound their
sum within `u64`, removing redundant overflow work while retaining all caller
limits and source/metadata extent checks. Output-bound planning uses constant
division for ordinary 64-KiB frames and still honors smaller configured
chunks. Independent source review found no lost validation or fidelity rule.

The full seven-crate consumer rerun passes 5,088 tests across 220 targets,
with seven existing ignored tests. This includes the earlier exact-budget
caption failure, malformed Snappy input and every new framing resource
dimension. The second pair uses fresh binaries from identical locks and a
verified 459-file candidate snapshot. Native results, allocation results and
final glossary/documentation receipts are reviewed below. The remaining
engineering work listed in 0361d stays required.

The [second receipt](results/0361d/native-v2.json) retains another five blocks
per route. All bytes and operation allocation metrics agree. Framing is
0.981× base and small catalog copy is 0.809×, but small decoding is consistently
1.002–1.007× (median 1.005×). This attempt is not accepted. Small compression
and sink-output measurements also need final qualification.

The next source repair preserves preflight error ordering while removing
arithmetic that the private constructor ceilings already prove safe on a
32-bit target. The encoded end of the current frame is the cumulative
compressed-byte count, so one value serves both range and stream admission.
Header ends, frame counts and decoded-length sums stay below their integer
ceilings. A new test checks aggregate observations and the ordering of frame,
chunk, stream, marker and truncation failures. Independent source review
found the arithmetic and ordering equivalent.

Compression also removes the repeated frame reservation after its initial
checked reservation. Each admitted encoded frame fits both the raw Snappy
bound and the configured chunk ceiling; each admitted prefix fits the stream
ceiling. Those are exactly the bounds used by the initial reservation, so an
accepted append cannot allocate again. Actual encoded-size checks remain
before appending. Independent source review confirmed the reservation proof,
including tiny ceilings, partial final chunks and multi-frame failures.

Both substrate crates pass all 313 unit, integration and documentation tests
after these final private changes, with warnings-denied Clippy also passing.
The [third comparison](results/0361d/native-v3.json) preserves five paired
blocks per route. Every leg exceeds 431 ms, all output checks pass, and
operation allocation metrics are unchanged. Small decoding is 0.996× base,
canonical framing 0.964× and small catalog copying 0.820×. Small compression
is still variable (0.950–1.058×, median 1.027×); large compression and sink
output have medians of 1.006× and 1.005×. These three routes require a longer
qualification before acceptance: nine paired blocks with one-second target
spans on CPU 29, using these same protected binaries. All original samples
remain part of the record.

The [longer qualification](results/0361d/native-v3-followup.json) rejects
small compression: eight of nine paired blocks are slower, with a 1.043×
median. Large compression is 1.004× and large sink output is highly variable
(0.907–1.099×). These observations are retained rather than accepted as new
exceptions.

The next repair lets the raw encoder write directly into the already-reserved
output whenever the complete conservative frame workspace fits. Tight output
ceilings keep the temporary-frame path: a compressible result can fit even
when its conservative raw workspace does not. The same encoder, chunk
geometry, actual encoded ceilings and frame headers apply to both paths.
Ordinary output avoids one allocation and payload copy per frame. A new
test compares direct and tightly capped multi-frame output byte for byte,
round-trips it and checks the one-byte-short refusal. Independent source
review confirms the reservation, 32-bit arithmetic, byte layout and failure
ordering. Family ledgers conservatively retain the temporary-frame allowance
needed by tightly capped profiles. All 314 substrate tests and warnings-denied
Clippy pass after a duplicate test import was corrected; both compile attempts
are retained.

## Final comparison

The [fourth comparison](results/0361d/native-v4.json) uses fresh release
binaries, identical locks and the same probe. All 459 candidate source files
match the checkout before timing. Five paired ABBA blocks per route target
500 ms per leg on CPU 0; the shortest measured leg is 423.89 ms. Separate
instrumented binaries report allocation metrics. All byte/checksum, framing,
round-trip and exact-sink checks pass.

| Operation | Median candidate/base |
| --- | ---: |
| Compress 1 KiB | 0.951079 |
| Compress repeating 4 MiB | 1.003367 |
| Compress incompressible 4 MiB | 0.681724 |
| Decompress 1 KiB | 0.999270 |
| Decompress repeating 4 MiB | 1.000319 |
| Decompress incompressible 4 MiB | 1.001987 |
| Calculate 4-MiB output bound | 0.170372 |
| Validate 512-object canonical framing | 0.978911 |
| Copy 1-KiB catalog | 0.819333 |
| Copy 4-MiB catalog | 1.011455 |
| Write 1-KiB catalog to sink | 1.000730 |
| Write 4-MiB catalog to sink | 1.001771 |

The large-copy result receives a separate
[nine-block qualification](results/0361d/native-v4-copy-followup.json) using
the same binaries and one-second target spans. Its shortest leg is 962.36 ms;
the median is 1.002398×, with paired blocks from 0.993512× to 1.039342× and
four of nine below baseline. Both runs remain in the record. This follow-up
does not reproduce a consistent large-copy regression. The other near-base
routes likewise have mixed paired blocks; these measurements support the
repair for the tested workloads, not exact timing equality for every input.

| Compression input | Peak additional bytes, before → after | Total requested bytes, before → after | Allocation calls, before → after |
| --- | ---: | ---: | ---: |
| 1 KiB | 2,456 → 1,230 | 2,456 → 1,230 | 2 → 1 |
| Repeating 4 MiB | 5,004,874 → 4,928,384 | 9,823,744 → 4,928,384 | 66 → 2 |
| Incompressible 4 MiB | 5,004,874 → 4,928,384 | 9,823,744 → 4,928,384 | 66 → 2 |

All other operation allocation counts and byte metrics are unchanged. The
reported peak is incremental over live fixtures and the warm result, not
process RSS. Additional retained bytes after each operation are zero on both
sides. Tight encoded ceilings still use the conservative temporary-frame
allowance already charged by the family ledgers.

Coordinator review accepts the repair without adding a performance
exception. The rejected first three candidates and both longer qualifications
remain reproducible through the committed probe sources, commands, source
digests, binary hashes and external artifact sizes/hashes. The full consumer
suite preceded the final private Snappy changes; the final 314 substrate tests
and Clippy cover those changes. W4/W4b/W5/W8 integration and the final full
gate remain required.
