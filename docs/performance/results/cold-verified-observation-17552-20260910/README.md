# Verified page-cache observation at 17552b1b4

This capture validates the opt-in cache-evidence path for `opc_file_source_open` on the recorded host and Rust 1.95 release build. Independent review passed 82 checks. Five cold samples each show zero resident, dirty, and writeback bytes before the operation, valid matching post-operation `fincore` provenance, and 81,920 process read bytes. Five warm and five cold samples ran in ten distinct children. No payload parts were materialized.

The warm and cold vectors are **not a matched latency comparison**. The verifier pads the ZIP from 16,783,632 to 16,785,408 bytes; logical reads are 1,040 bytes for warm and 66,576 for cold. The raw timings and recomputed Student-t intervals remain diagnostic evidence only. This does not establish physical-media temperature, device-cache state, a speedup, or completion of the performance program.

Source commit: `17552b1b41796f8c0e3953017931f497b1837b3b`. Binary SHA-256: `bb60a945eb855b3b87fec50765c3a50b73fea39c508b49751777acabd3bae57d`. The build receipt binds the toolchain, lock, source, host and build identities. CPU 2 was released after capture.

Each gzip file preserves the exact original bytes. `publication.json` binds both raw and compressed hashes. The independent receipt contains the checks and statistical replay. Reproduction commands are in `BUILD_PREPARATION.md.gz` and `root-capture.json.gz`; the latter records argv but omits the launcher's TMPDIR environment assignment. The additive publication note records that limitation without changing the original evidence.
