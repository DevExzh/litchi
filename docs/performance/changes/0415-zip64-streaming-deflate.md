# Change 0415: ZIP64 streaming Deflate

This change establishes valid ZIP64 framing before an unknown-size entry is
written. It is a capability change, with descriptive memory observations and
ordinary ZIP32 regression guards; it makes no speedup claim.

## Implementation

The low-level streaming builder accepts `.zip64(true)`, and the consuming
writer exposes `start_file_owned_zip64`. Both select a version-45 local header,
two fixed size sentinels and a ZIP64 extra containing two zero placeholders.
The selected mode persists through finalization: even a small final entry uses
a signed 24-byte descriptor and explicit central ZIP64 size fields. Ordinary
ZIP32 mode refuses an unexpected size overflow before writing a descriptor.
Counters use checked arithmetic; generated central metadata is prepared before
descriptor output, and conflicting caller-provided ZIP64 extras are rejected.

The Office transport selects this mode from its existing admitted per-entry
and compressed-byte limits. Default limits preserve ordinary ZIP32 output.
Raised limits permit ZIP64 counts, offsets and sizes while retaining configured
metadata, payload and output ceilings, including generated central ZIP64 extras.
Caller-owned non-atomic sinks retain byte-progress and poisoning semantics.

Strict input validation accepts the corresponding streaming placeholder form
only with bit 3, version 45 or later, both sentinels and two zero values in a
valid ZIP64 size extra. Final descriptor CRC/sizes and central metadata still
must agree. Nonzero mismatched local sizes remain errors. This also admits the
independent large Python producer used by the interoperability check.

Preservation regeneration selects ZIP64 from a checked conservative bound
before starting Deflate and removes its previous ZIP64 Deflate refusal. That
helper still owns buffered input and output; it is separate from bounded
streaming creation and does not gain a bounded-memory regeneration guarantee.

The framing follows [PKWARE APPNOTE sections 4.3.9, 4.4 and 4.5.3](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT).
For interoperability, central ZIP64 size fields are explicit even for small
opt-in entries. Python's small `force_zip64` variant with ordinary central size
fields remains an input-compatibility gap requiring a separate descriptor-width
decision.

## Evidence and scope

The [retained bundle](../results/change-0415/) contains the deterministic Python
producer, Rust probe, protocol, raw guard samples, independent readback and
verification logs. The control is `916f60c38`; production candidate is
`caf0d9394`. The final integrated run passes 1,222 tests. Both additional
large/interoperability tests pass when explicitly selected; 43 doctests pass.
Warning-denied rustdoc, scoped Clippy, changed-file formatting, boundaries,
coverage, eight strict claims and report classification pass. The existing
unexempted Clippy and unrelated DOCX formatting qualifications remain recorded.
Both existing ZIP/OPC fuzz targets complete 1,000 coverage-instrumented
AddressSanitizer iterations without a finding.

The independent corpus is a four-member generic OPC package. Its `large.bin`
contains exactly 4,294,967,296 zero bytes. Python writes to a sequential sink
and verifies all members through a 64 KiB buffer. The OPC integration check
edits a small XML part while checking the large member's raw local and central
preservation, including the permitted offset relocation.

The Python source uses per-entry ZIP64 size metadata with an ordinary global
tail: its central directory remains small. Candidate-produced large streams
also carry the explicit ZIP64 global tail. Both forms are validated without
conflating entry metadata with tail framing.

## Resource observations

Each zero-input observation is a fresh, CPU-2 process writing a caller-owned
file on `/tmp` tmpfs. Full Rust and Python verification run separately. These
are single observations, with no confidence interval or durable-device I/O
claim. RSS includes the writer process and excludes output storage/page-cache
memory.

| Logical zero input | Borrowed writer time | Owned writer time | Borrowed peak RSS | Owned peak RSS | Archive bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| 64 MiB | 20.080 ms | 19.586 ms | 2,668 KiB | 2,680 KiB | 65,374 |
| 256 MiB | 80.947 ms | 77.303 ms | 2,688 KiB | 2,684 KiB | 261,059 |
| 4 GiB + 1 byte | 1,282.681 ms | 1,238.408 ms | 2,688 KiB | 2,688 KiB | 4,174,773 |

A separate incompressible-output check writes 4,294,967,297 bytes from a fixed
64 KiB pseudo-random block. The actual Deflate payload is 4,296,277,865 bytes;
the complete archive is 4,296,278,125 bytes. Both size fields cross the ZIP32
boundary. The owned writer and subsequent bounded Rust readback use 2,688 KiB
peak RSS. Writing takes 57.217 seconds including the input-CRC observer; Rust
readback takes 0.490 seconds. The writer timer starts after file creation and
input-block generation; process RSS includes that setup. Independent Python
verification confirms the
complete plaintext count, CRC and SHA-256. This closes the physical-output gap
left by highly compressible zeros for this measured transport path, while
remaining a tmpfs observation rather than a native storage benchmark.

The [CPU flame graph](../results/change-0415/profile/flamegraph.svg) is a
separate owned zero-stream diagnostic with 1,006 samples and zero lost samples.
`longest_match` accounts for 48.05% of self samples and `slide_hash_chain` for
20.99%; Deflate work dominates. It is not a paired CPU improvement measurement
and does not profile the incompressible-output case.

## Ordinary ZIP32 guards and disposition

Eight rows cover borrowed/owned transport, zero/pseudo-random input, and
16 KiB/1 MiB payloads. Both ABBA captures retain all individual p50/p95/p99,
RSS values, raw samples, output-byte counts, write counts and output CRCs.
Output observations match across all four legs of every row.

The initial 300-sample capture triggers review for:

- borrowed zero 16 KiB RSS: +9.65%;
- borrowed pseudo-random 1 MiB p50/p95/p99: +7.63/+7.77/+8.13%, and RSS +8.18%;
- owned zero 16 KiB p99: +38.32%.

The 1,000-sample follow-up uses the same binaries and retains new triggers:
borrowed zero 16 KiB p99 rises +27.64% and +7.40% in the two pairs, and borrowed
zero 1 MiB RSS rises +9.90%. Follow-up pseudo-random 1 MiB median costs are
+2.39% to +2.74%, with p95/p99 below the 5% trigger. The original and follow-up
summaries also retain within-revision drift separately. RSS crossings are
220–264 KiB differences between low-MiB guard processes; they remain review
items, alongside the unresolved small-entry p99 increases.

Retain this as a necessary framing, resource-accounting and compatibility
capability. The measured common-path cost and unresolved tail/RSS triggers are
explicit; this is not a clean latency-regression pass and no speedup is claimed.

These checks concern ZIP transport streaming creation and a source-backed Part
edit. They do not establish semantic row, paragraph or slide streaming, the
four distinct append scenarios, cold/remote/native behavior, concurrency,
worker scaling, or the complete non-iWork CRUD program.
