# OLE2 stream validation

This frozen batch covers MS-OLEDS 2.3.4–2.3.6 codecs and their target-selected OLE object integration. The receipt identifies the baseline and exact candidate source hashes.

The combined CFB/OLE gate passed 507 unit/integration tests and 14 doctests, with one existing doctest ignored. Strict all-target Clippy and changed-file formatting checks passed. Regressions cover malformed and bounded fields, independent presentation-count and TOC-count limits, signed dimensions, source sharing, exact no-ops/inverse, stale-patch refusal, atomic publication and directory metadata preservation.

Payloads remain opaque and inert. This evidence does not establish OLE1 support, native-producer interoperability, throughput, or total memory use. Compressed logs are stored with deterministic gzip headers.
