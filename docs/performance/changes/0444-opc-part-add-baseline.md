# 0444: Source-backed OPC Part-addition observed baseline

Add the opt-in `opc_part_add_lifecycle` benchmark and an explicit fixture exporter.
A caller-owned source with 64/1024/4096 Parts gains one 64 KiB typed leaf and one
root internal relationship through the existing sequential topology publisher.
This is synthetic OPC package topology, with a binary officeDocument target;
it does not create a semantic Office document owner.

The frozen baseline contains 12 reports/360 samples, normal and allocator modes,
two reversed repeats, CPU 2 and one worker. Normal p50 is 1.101–1.108,
8.017–8.020 and 61.689–61.876 ms. No repeat check exceeds 5%. Allocation calls
are 3,526 / 51,538 / 205,145; operation peak above entry is 705,777 / 2,716,017 /
9,181,553 bytes. Endpoint live delta and retained sink output are zero. All
[rows and uncertainty](../results/change-0444/measurements.md) remain visible.

The source observer accounts for 55.24% of whole-process sampled self time and
scans member ranges on every read. Timings therefore describe the observed path;
a matched plain-source baseline is required before attributing size growth to
production topology. This batch removes no production work and claims no
speedup, bounded total memory, native compatibility, cold/range or scaling gain.

Verification: 372 harness tests and 458 OPC tests passed, each suite with one
existing ignored test. Workspace/harness checks, warning-denied rustdoc, explicit
formatting and boundary checks pass. Strict Clippy retains inherited debt with
zero new diagnostics. Independent ZIP/XML/raw-record/payload checks pass on all
six exported archives; 29 corruption probes reject. All failed development and
pilot attempts are retained. The real protocol freeze timestamp precedes every
retained capture; implementation, fixture validation and pilots precede it.

The reusable source/sink metric envelope aligns actual reads to elapsed/sample
order without inventing materialization or codec counters. No production API,
dependency, unsafe code, scheduler or runtime I/O changed. The selector registry
is 437, with 36 defaults; semantic representative coverage is unchanged.
The full non-iWork goal remains active. Evidence and replay instructions are in
[the bundle](../results/change-0444/README.md).
