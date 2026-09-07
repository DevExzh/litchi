# Matched timing and allocator protocol

The final baseline and candidate use identical standalone harness sources. Only
the PPTX production and focused test files differ. Both normal and allocator
executables are retained with exact source manifests and binary hashes; the
revision alone cannot distinguish these working-tree builds.

Freeze protocol.json before pilots. Run 12 one-sample controls, then 24 fresh
processes with three warmups and thirty samples each on CPU 2, one worker.
Sixteen normal reports cover bytes/range and plain/media-rich workloads in
forward/reverse order (480 samples). Eight allocator reports separately cover
bytes and both corpora in forward/reverse order (240 samples). Allocator timings
are not pooled with ordinary timings. Range uses the existing 64 KiB cap,
200 microseconds per call, 25 MiB/s, separate-sleep simulation.

The media-rich allocation gate requires both repeats to remove at least
16,773,120 planning allocated bytes and planning live-growth bytes. Retain all
absolute paired and repeat changes over 5% for individual review, including
improvements. Phase medians, complete API p50/p95/p99, conditional bootstrap
median intervals and process RSS remain available in raw and derived data.

Allocator regions cover each open, planning and publication. Their live peaks
include region-entry live allocations and are not RSS. Process RSS includes
untimed deterministic fixture construction. Neither metric certifies cold I/O,
physical networking, native Office application behavior or scaling.

Reproduction drivers: rebuild-baseline.py restores the candidate after building
the baseline; run-checks.py builds the candidate after validation; run-fuzz.py
runs the unchanged OPC substrate fuzz target; freeze.py and run-matrix.py capture
the matched matrix. These capture drivers require the repository and original
exclusive scratch paths. verify.py replays the exported evidence independently
without source checkout, build products or scratch paths.

After the primary matrix, its two plain/bytes p99 triggers motivated a separate
fixed investigation recorded in confirmation-protocol.json: two ABBA blocks,
eight fresh processes and 240 samples, using the same binaries, CPU and warmups.
The copied primary media-allocation thresholds do not apply to this plain-only
investigation. run-confirmation.py records the exact commands; per-process and
combined descriptive distributions stay separate from the primary matrix.
