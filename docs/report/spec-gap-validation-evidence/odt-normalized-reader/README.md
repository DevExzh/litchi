# ODT normalized namespace reader migration

The ten ODT source consumers listed in `source.json` now use the shared normalized namespace reader and compare resolved URIs without decoding them again. Source events and positions remain borrowed and unchanged. This admits escaped namespace bindings while retaining source-preserving editing behavior.

Root validation used the exact combined capture in `combined-validation-source.json`: 2,230 passing tests across 127 result groups, one ignored test; all-target Clippy with warnings denied; rustdoc with warnings denied. Commands and exit codes are retained beside raw logs. Cargo.lock is the validation checkout lock. The combined capture includes uncommitted ODS work; these receipts do not approve that work. The ODT files and common dependency hashes were checked against the live tree before this commit.

Independent review approved the listed ten ODT source files and test, with 615 ODT tests and 13 shared reader tests passing independently. It checked raw and buffered event preservation, single namespace normalization, limits, cancellation, exact no-ops and atomic failure behavior. Other ODT namespace-reader consumers and global namespace coverage remain outside this batch. No runtime speedup or allocator claim is made.
