# 0842 selection and interpretation

The preceding 0841 dependency audit identified a current-lock configuration
regression: package-scoped OPC resolves flate2 1.1.10 with the zlib-rs backend,
but disables defaults without selecting the new runtime_detection feature.
Historical 0836 resolved flate2 1.1.9, which requested zlib-rs/std directly.
Its timings cannot be reused as this experiment's baseline.

The candidate changes the workspace and independently declared soapberry-zip
flate2 dependencies to require at least 1.1.10 and explicitly request
runtime_detection. The root versions stay flate2 1.1.10 and zlib-rs 0.6.7.
The live perf-baseline lock must also move from 1.1.9 to 1.1.10 so its locked
build remains usable. Unused miniz/adler lock entries are removed by Cargo;
no other retained package version changes. The perf-execution and native-resave
locks also gain the crc32fast dependency edge. Native-resave additionally has
baseline-confirmed stale local dependency edges: Cargo refreshes crypto, DOC,
ODT and XLS edges and adds same-file/winapi-util, identically with the original
compression declarations. No existing registry version/checksum changes.
These two standalone lock edits follow the measured candidate build and are
not inputs to the probe or root workspace; source-final.json and
native-resave-lock-drift.json record that delta and its baseline control.
All historical packets stay intact.

The standalone probe intentionally omits other feature consumers. Its retained
metadata proves that the only resolved feature differences are
flate2/runtime_detection and zlib-rs/std. Both legs have identical registry
versions/checksums. Cargo adds an existing crc32fast lock dependency edge when
the feature is activated; both stage-specific locks are retained.

## Scope and fixed protocol

Four opaque OPC parts use Office-like member names but do not form a semantic
DOCX, XLSX, or PPTX document. Tiny has four 64-byte parts. Each other shape has
three 1 MiB valid synthetic XML parts and one 1 MiB opaque part; that last part
is pseudorandom only in random, and a repeated byte pattern in xml/repeated.
Their different XML shape labels also change the XML payloads. Six ZIP members
include content types and the empty package relationships member.

Write times PackageWriter::to_bytes on a prepared package. Read times ZIP
reader construction and all six member reads, retaining decompressed payloads
until the timer stops. Changed times a source-backed open, topology replacement
plan for custom/payload.bin, and publication; input and replacement clones are
prepared before its timer. Expected archive construction, fixture generation,
hashes, validation, and result destruction stay outside these clocks.

These are warm processes with validation interleaved between samples. Write
and changed decode their output twice after stopping the clock: first for the
logical digest, then for independent expected-payload checks. Read validates
its returned payloads. Both binaries follow identical code, but the feature
also accelerates untimed decoding; cache/thermal/frequency state is therefore
not isolated from that work. Results describe this precise protocol and do not
establish native cycle shares or a pure compressor-only causal decomposition.
No validation cost is subtracted from elapsed values.

Six blocks alternate leg and case order, CPU 12, 30 samples after three
warmups. All cases and all tails remain visible. /usr/bin/time RSS covers the
whole child including prepared fixtures and validation, not an operation heap
peak. No allocation, physical-copy, cold-storage, remote-range, scaling, or
native Office compatibility promotion follows from this packet.

## Admission and decisions

Both source states must qualify exact archive bytes, independent Python logical
payload/CRC checks, and retained compressed payloads. Changed outputs must
preserve untouched local ZIP spans and central metadata apart from relocated
offsets. The auxiliary codec matrix covers whole borrowed payloads, owned
writes in chunks of 1/7/32768 bytes, sized writes, two reader-fed APIs, repeated
writes, and preservation copy. Nine named outputs include a duplicate
borrowed/reusable control, so they are not nine independent algorithms.

The analyzer's pass means evidence is internally valid. A separate final
decision reviews all p50/RSS guard flags, all increasing tail point estimates,
and block spread flags. Ratios with zero denominators (such as all-zero major
faults) are undefined, never invented as 1. p50 and RSS must be positive.
No candidate is adopted solely because its evidence parses.

| Constraint | Candidate boundary |
| --- | --- |
| ADR 0001, 0006 | Exact output and typed-error suites precede speed claims. |
| ADR 0002, 0011, 0024 | Dependency declarations stay with existing owners; no archive API moves. |
| ADR 0005 | No cache, I/O, preservation, durability or parallelism policy changes. |
| ADR 0003 | No snapshot, commit, patch or inverse contract changes. |
| GOAL SIMD requirements | Existing dependency CPU detection and scalar fallbacks; no new unsafe production code. Host evidence is x86_64 only. |

Full workspace all-target and documentation gates are blocked by baseline-confirmed
iWork defects (example return type and two broken rustdoc links). iWork is
excluded by the user. Workspace library/test checks and library Clippy plus
all-target checks, documentation, and tests for the five affected crates form
the applicable gates; the original failures remain visible in commands/.
