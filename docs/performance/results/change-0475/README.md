# PPTX streaming CPU and allocation attribution (0475)

This bundle profiles the unchanged public `pptx_streaming_create` large case:
8,192 plain-text slides, 16,421 ZIP members and a 7,940,406-byte output. Its
purpose is to explain the allocation work measured in 0474 before choosing a
shared transport change. It makes no production optimization, before/after,
registered latency or constant-memory claim.

The original normal executable was reused only after its complete SHA-256 and
size matched the sealed 0474 build. The original source path was recreated at
`60a1a4300a2844f372c0c916b5be69c50ce38668`; all 7,034 source hashes, both ignored
fixtures, clean Git identity and the executable are checked before and after
captures. `reuse-validation.json` records successful verification of the full
0474 bundle. `reuse/` holds selected authenticated artifacts and the old seal;
it is deliberately not a duplicate of that entire bundle. No Rust source or
dependency changed in 0475, and no new release build is claimed.

`protocol.json` freezes the driver hash, binary reuse, corpus, CPU 2, one worker
and seven-lane order before preparation. Normal runs bracket `perf stat`, two
CPU recordings and two Heaptrack recordings. Normal/counter/CPU lanes each
use three warmups and thirty samples; each heap lane uses zero warmups and
one sample. Every report independently checks the same preflight/output oracles.
All owned heavy commands use `/tmp/litchi-goal-0475/cpu.lock`.

Profiler captures include the whole process: corpus construction, semantic
reopen, timed writer, observers and teardown. The exact `run` ancestry matters
for writer attribution. CPU periods are sampled `cycles:u` weights, not elapsed
phase durations. Heap allocation calls, requested bytes and event-order live
peaks have distinct meanings. Process RSS and profiled timing remain separate
from the normal operation and the prior allocator-observer baseline.

Capture receipts bind raw artifacts. `compression.json` connects deterministically
compressed exports to the hashes and byte counts of their removed originals.
Original Heaptrack zstd traces and compressed `perf.data` are retained alongside
portable text exports, exact export commands and stderr. Parsers and tests live
within this directory, so replay does not need the source checkout, executable,
an earlier bundle, perf, Heaptrack or network access.

After all exports completed, `split.py` stored each original Heaptrack trace
in ordered 64 MiB parts. `chunking.json` binds each part and the complete
concatenation to its original capture hash and length. The original logical
file can be restored by concatenating its listed parts in order. Portable
verification checks those bytes directly; it does not require restoration.
The decoded gzip exports remain available for the portable heap parser.

The original demangled Heaptrack filter returned no stacks because the trace
retains Rust v0 mangling. That result remains unchanged. `supplement-protocol.json`
declares the exact captured run token, and `supplement.py` exports separately
named filtered stacks from the same traces. `supplement-compression.json`
binds those additional compressed exports.

See `summary.json` for complete normal/profiler vectors and counter context,
`cpu-P1-attribution.json` and `cpu-P2-attribution.json` for sampled callchains,
and `heap-attribution.json` for requested-byte and lifetime attribution.
`owners.json` derives the disjoint backend initialization and flate2 buffer
totals from exact stack markers: 6,778,588,800 requested bytes in both traces,
99.543891% of requests under `run`. This includes the broader run context,
not just the timed operation. Generic nearest-frame category names are not
sufficient to distinguish buffer construction from ZIP metadata.
`transport-audit.md` records the candidate mechanism and correctness obligations;
`heap-format-notes.md` identifies the upstream event format and parser limits.

Portable verification, after copying this entire directory:

```sh
python3 -B verify.py
python3 -B analyze.py --check
python3 -B -m unittest discover -s . -p 'test_*.py'
python3 -B replay.py
```

`verify.py` checks custody, report arithmetic and attribution input bindings.
`replay.py` additionally recomputes both CPU and heap attribution outputs
exactly. Heap replay scans the multi-gigabyte decoded streams and takes longer
than the other portable checks. Its default projection visits only exact
writer-run allocation lifetimes; whole-process requested-byte totals and
timeline peaks remain unavailable in that projection.

The seal covers every regular bundle file except itself. Validation receipts
retain each actual command, source hashes, outputs and exit status. Temporary
checkouts and binaries are not needed for portable verification; `--live`
additionally checks them while they exist.

The full non-iWork goal remains open. Reducing repeated compressor allocation
work would still leave growing OPC/ZIP name indexes and directory metadata.
Fresh creation also remains separate from logical append, Part addition and
arbitrary edits followed by repackaging.

Final sealed live verification and post-cleanup standalone verification pass.
The standalone copy reproduces both CPU and heap attribution outputs exactly,
passes all 33 tests and rejects an altered owner-total output. The copy is
removed. The original frozen captures and exports remain unchanged; the
interrupted analysis and initial stale-scope verifier failure remain retained.
