# OPC physical entry admission V5

Source-backed and metadata-only package admission now enforce ArchiveTotalEntries against the physical ZIP count before indexed-directory allocation. This includes directory records, matching eager ingress. The underlying ZIP index bounds allocation by the declaration and rejects underreported counts before retaining extra records.

The independent review reproduced the four-file patch, checked all ingress paths and underreported EOCD rejection, and verified all 24 resource kinds with exact/over/under checks. Smoke coverage now includes a nested XML fixture, directory entries distinct from ordinary file members, strict invalid-limit builder assertions, and byte-exact replacement/addition readback through source, eager and raw ZIP paths.

Root independently built a clean archive of 6bba8c3d48fc7d360c50f2fda607afc173e10d98 with the four reviewed files and explicit validation locks. Root gates passed 551 tests across 17 targets including doctests, one existing ignored test, strict all-target/all-feature Clippy, warnings-denied rustdoc, scoped Rust 1.95 formatting and the bounded smoke. Source and locks remained unchanged. Root verified current owned preimages and preserved the pending OPC README.

The 10,000-execution instrumented campaign published in opc-fuzz-v3 remains historical evidence for that earlier source. V5 has new package/smoke/review evidence; no new instrumented campaign or later-source 10k claim is made. The author separately reused a prior RC4 integration directory in error; this publication uses only root's independently constructed clean candidate and the original immutable V5 source/review evidence, not that reused directory.
