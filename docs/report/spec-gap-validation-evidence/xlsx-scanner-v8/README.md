# XLSX sparse verification scanner V8

The edit transaction retains only changed coordinates during verification. A narrow borrowed XML scanner handles eligible scalar worksheets; unsupported shapes use the existing MCE scanner. Malformed input and resource failures remain typed errors.

The exact four-file delta was validated on clean integration base de80c38248ff797501564886c5df0987088f430c: 1,437 tests across 63 targets passed, none ignored, with strict all-target/all-feature Clippy and rustdoc. Source hashes before and after gates match. The root build lock is retained separately (SHA-256 8248c6b7787a9726e9f99396a5de147bc2f68acf0afef44b710a6392f940bd88); offline reconciliation changed local XLDM package/dependency edges only, with no registry changes. Source preimages were rechecked against the publication HEAD before integration.

Performance evidence is historical: the matched pair shares base 21d5fb0bef3a36ee42903e5cc7cb2fd98373150e. Forty normal samples per side yielded median -8.460474%, nearest-rank p95 -7.868794%, and p99 -7.515784%. Operation-scoped allocator calls fell 20.897961% and allocated bytes 32.814253%; reallocations rose 0.315610%. Allocator elapsed time is excluded from latency. These numbers do not measure the later integration head or a combined gain with the earlier plain-cell optimization.

The portable archive retains 59 mapped artifacts, including all 31 released raw outputs, matched source identities, locks, capture plan, reviews, and corrected historical analysis. Large release binaries are omitted with their exact identities recorded. Root extracted the archive afresh, ran its verifier without executing any workload, recomputed the statistics, and verified the 60 extracted files remained unchanged.

To verify, extract historical-portable-bundle.tar.zst and run verifier/verify-v8-portable.py with --bundle-root pointing to the extracted bundle and --output-dir pointing to a separate new directory. Raw root logs and the build lock are gzip-compressed without changing their bytes.
