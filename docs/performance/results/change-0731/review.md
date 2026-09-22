# 0731 independent profile review

This review covers the nine retained PPT processes for `ppt45543`: three
native processes with 50 samples and three warmups each, three allocator
processes with one sample and no warmup, and three Callgrind processes with
one sample and no warmup. All processes ran on CPU 12 against the warm-cache
`45543.ppt` fixture. The probe made no production-source change.

The nine JSON reports reproduce the sealed 0728 PPT format oracle. Every
report has the same source digest, expected output digest, stream inventories,
changed-length proof, and semantic witness. Every requested sample has an
exact output match and a passing oracle. The eight direct corruption controls
(`missing_stream`, untouched-stream mutation, root and stream metadata
mutations, source swap, wrong slide identity, and survivor-payload mutation)
are rejected in every report. This establishes the correctness and custody
requirements for this one measured workflow.

The native timing context is stable across the three independent processes:

| repeat | p50 (ns) | mean (ns) | p95 (ns) | p99 (ns) | maximum (ns) |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 1,070,685.5 | 1,058,066.46 | 1,178,816 | 1,192,495 | 1,192,495 |
| 1 | 1,067,220.5 | 1,058,567.18 | 1,179,716 | 1,219,416 | 1,219,416 |
| 2 | 1,058,430.5 | 1,054,805.16 | 1,168,556 | 1,181,686 | 1,181,686 |

Relative to repeat 0, the repeat 1 p50 and mean changes are −0.324% and
0.047%, and repeat 2 changes are −1.145% and −0.308%. Neither repeat crosses
the registered 5% review flag. This is a warm-cache context measurement, not
a before/after optimization result.

The allocator process produced the same boundary-relative region in all three
runs: 11,776,674 allocated bytes, 11,386,530 deallocated bytes, 5,663
allocation calls, 2,686,521 peak live bytes, and 390,144 retained bytes. The
region retains the returned output for later validation; these counters are
not RSS and the three one-sample processes do not establish a memory
distribution.

The raw Callgrind files contain only the `Ir` event. Their totals are
59,578,167, 59,578,131, and 59,578,178 instructions for repeats 0, 1, and 2.
In each file, the collection owner
`ole_format_save_probe_0731::measured_public_format` has one self instruction
and one `public_format_edit` child edge carrying respectively 59,578,166,
59,578,130, and 59,578,177 instructions. The owner inclusive totals exactly
equal the file totals. The parent edges are one call each from
`timed_format` to the owner and from `run` to `timed_format`, with the same
inclusive totals. The repeated one-call edge and exact total equality confirm
that collection was toggled around the intended owner invocation. The equal
costs shown for its ancestors by `callgrind_annotate` are inclusive call-graph
propagation; they are not evidence that collection was enabled during setup,
oracle construction, or output validation.

The largest self entry is
`sha2::sha256::soft::unroll::compress`, 40,370,752 simulated instructions in
each profile, or 67.76097%–67.76102% of the total. `__memcpy_avx_unaligned_erms`
and `__memset_avx2_unaligned_erms` are the next two self entries at 14.94337%
and 9.81599%. These percentages describe the software-instrumented Callgrind
run only. They are not native fractions, latency fractions, hardware-counter
measurements, or optimization targets by themselves.

The post-profile feature witness records native `sha=true`, `avx2=true`, and
`avx512f=true`, while `valgrind --tool=none` reports `sha=false`, `avx2=true`,
and `avx512f=false`. The checked `sha2` 0.11.0 source dispatches to its x86
SHA implementation when the runtime SHA/SSE feature set is available and
otherwise falls back to the software implementation. The soft SHA symbol in
Callgrind is therefore a simulator-backend observation. It can identify that
the workflow performs repeated artifact hashing and justify a separately
controlled native investigation, but it cannot establish the cost or benefit
of changing that path on the native host.

No production optimization or adoption claim follows from this packet. A
follow-up native attribution should preserve the same exact PPT oracle and
use observer controls or host hardware counters with the native feature
dispatch intact. Any resulting claim must remain scoped to this fixture and
workflow until additional PPT corpus coverage is measured. This packet also
does not measure cold-cache behavior, RSS, I/O, concurrency, other PPT
fixtures, or any DOC/PPTX route.
