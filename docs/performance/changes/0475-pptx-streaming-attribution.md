# Change 0475: PPTX streaming CPU and allocation attribution

`performance_claim: none; descriptive profiling`

`claim_authorized: false`

This batch profiles the unchanged public `pptx_streaming_create` large case
to distinguish repeated allocation work from the structural memory growth
measured in 0474. There is no production or benchmark Rust change. The reused
normal executable matches the sealed 0474 binary byte for byte, and the clean
source checkout matches all 7,034 source hashes at
`60a1a4300a2844f372c0c916b5be69c50ce38668`.

All seven frozen captures pass their report and deterministic output gates:
two normal runs, one counter run, two CPU recordings and two Heaptrack
recordings. The five non-heap lanes each contain thirty samples after
three warmups; each heap lane has one sample and zero warmups. The case writes
8,192 slides, 16,421 ZIP members and exactly 7,940,406 bytes to the non-seek
hashing discard sink, which retains zero output bytes.

Normal p50 is 258.296564 / 254.160741 ms in the bracketing repeats. Mean,
p50, p95 and p99 drift are respectively -1.617728%, -1.601192%, -1.045072%
and -0.770084%, all below the frozen 5% review threshold. These are unchanged
code repeats, not before/after improvement evidence. Profiler elapsed vectors
remain separately identified in the bundle.

Whole-process `perf stat` reports 157,634,913,764 user cycles and
690,745,026,904 user instructions, with hardware counters scheduled for
83% of their enabled interval. The counts include setup, preflight and
observers, and perf scales multiplexed values. LLC load misses are unsupported;
L1 data-cache load misses are reported as zero, which does not establish that
the workload has no such misses. Neither event is silently substituted.

Normal maximum process RSS is 82,724 / 82,692 KiB. CPU profiler process RSS
is 511,996 / 82,740 KiB; Heaptrack process RSS is 313,660 / 314,312 KiB.
These whole-process values include profiler/setup effects and cannot replace
the prior operation-local allocator peak of 8,875,092 bytes above entry.

The large whole-process Heaptrack counts include repeated archive-reader
construction in materialized preflight. Shared/inlined allocator symbols can
show misleading higher-level labels; attribution requires the complete caller
chain and exact benchmark `run` versus `build_corpus` ancestry. This setup
work lies outside the timed operation. Filtered Heaptrack peak summaries are
also not operation-local peaks.

The frozen Heaptrack export used a demangled `::run` filter, but the captured
symbols use Rust v0 mangling. Its empty result is preserved. A separately
declared supplemental export uses the exact captured
`21pptx_streaming_create3run` token against the same traces. It changes no
capture, binary, source, timing or corpus identity.

CPU callchains assign 23.983597% / 23.912740% of whole-process sampled
period weight to exact benchmark `run` ancestry and 75.174% / 75.294% to
materialized preflight; the remaining weight is unclassified. Both exports
have zero lost, unparsed or truncated samples, with unknown-frame coverage
reported separately. Inside the writer context, `zlib_rs::deflate::deflate`
appears in 51.521283% / 50.860751% of sampled weight and
`zlib_rs::deflate::init` in 5.673627% / 4.961299%. These inclusive rows must
not be summed or converted to elapsed phase timings. The leading writer leaf
is `__memset_avx512_unaligned_erms` at 29.873% / 29.416%; its name alone
cannot assign all zeroing work to encoder initialization.

Both heap traces report exactly 311,120 allocation events and 6,809,648,242
requested bytes under the benchmark `run` ancestry, with a projected live peak
of 8,875,780 bytes and zero live bytes at exit. This context includes observer
and report work outside the timed operation: it has 127 more allocation events
and 44,229 more requested bytes than the prior operation observer, and its peak
is 688 bytes higher. The two scopes must not be renamed interchangeable.

Exact initialization stacks contain 16,421 backend allocations of 380,032
bytes: 6,240,505,472 requested bytes. Separate flate2 output-buffer stacks
contain 16,421 allocations of 32,768 bytes: 538,083,328 requested bytes.
Together they account for 6,778,588,800 bytes, or 99.543891% of requested
bytes under `run`, identically in both traces. These are cumulative allocation
requests, not concurrently retained gigabytes. The generic nearest-frame
`ZIP-metadata` category includes the flate2 buffer because ZIP types occur in
its generic signature; `owners.json` instead uses the exact, disjoint backend
initialization and buffer-construction stack markers.

The scalable heap parser scans metadata and only exact-run allocation/free
events, retaining their lifetime order. Whole-process requested totals and
timeline peaks are unavailable from that projection. Unresolved trace metadata
and excluded preflight scope remain explicit. Category peaks are independent
lifetime projections and must not be summed. The first analysis attempt was
terminated after a slow descriptor filter; its receipt remains with exit -15.
Factoring the same exact-token filter passed exhaustive membership tests and
the replacement analysis. A first verifier attempt rejected a stale scope
field name; the corrected live verification passes with both attempts retained.

The [evidence bundle](../results/change-0475/README.md) retains original
compressed traces, portable exports, source/build reuse authentication, exact
commands, report vectors, parser tests and immutable attempt receipts. The
[transport audit](../results/change-0475/transport-audit.md) identifies private
raw compressor-state reuse as the next allocation-work candidate, subject to
measured attribution and a matched implementation comparison.

Compressor reuse must preserve independent raw Deflate members, complete
StreamEnd/descriptor ordering, CRC and accepted-input accounting, limits,
short-write behavior, poisoning, cancellation and deterministic output. It
would not remove growing OPC/ZIP name indexes or the ZIP central directory.
An explicit-window total-memory contract needs separate metadata design and
measurement. Fresh creation, logical append, Part addition, arbitrary
repackaging, native breadth and scaling remain separate open requirements.
The full non-iWork goal remains open.

Validation passes 33 Python parser/filter/evidence tests, exact normal and CPU
summary replay, the strict ten-claim registry and sealed live verification.
No Rust source changed; the 0474 Rust gates are reused with authenticated
source/binary identity rather than reported as freshly rerun. The standalone
post-cleanup replay record retains final CPU/heap recomputation and mutation
rejection status. Owned runtime checkouts and binaries are removed; shared
Cargo caches and the two user-owned files remain intact.
