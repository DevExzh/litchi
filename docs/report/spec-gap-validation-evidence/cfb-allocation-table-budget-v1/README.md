# CFB allocation-table budget validation

Validated the exact source files in `receipt.json` on isolated baseline
`1ffcebbf7`: 314 unit/integration tests and 14 doctests passed; one existing
doctest is ignored. Strict library/test Clippy, formatting and whitespace
checks passed. Logs preserve their exact original bytes in deterministic gzip.

The profile checks combined declared FAT/DIFAT/MiniFAT sector bytes and their
`u32` location vectors before reserving the physical-sector index or tables.
Tests cover exact and under-limit admission, invalid limits, hostile counts
with no table reads, shared limit propagation and revalidation, blocked
validation reports, and sparse v4 range-lock readback beyond 2 GiB.

The 64 MiB default and 2 GiB hard ceiling are resource policies, not format
maxima or measurements of total process memory. This does not bound directory
bytes, physical-sector roles, chain scratch or source memory, and makes no
throughput claim.
