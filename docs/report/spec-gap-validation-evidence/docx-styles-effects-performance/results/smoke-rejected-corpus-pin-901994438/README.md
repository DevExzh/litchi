# Rejected DOCX stylesWithEffects smoke attempt

This directory retains the complete raw output from the first d000 candidate
smoke attempt. The frozen harness checkout was
`901994438a3202fae2c3a13535f538c222a08240`, with production and runner source
pin `d000d977b99e03f8542c7dae74acf767a91b1feb`. The operator session was
`31272` and exited 1.

The runner completed all 52 lane processes and wrote 166 raw files: 52 JSON
receipts, 52 `/usr/bin/time -v` sidecars, 52 empty stderr sidecars, and the
metadata, build, manifest, command, and provenance files. The raw files total
1,771,338 bytes and are listed with SHA-256 digests in `retained-files.json`.
Their receipts carry the d000 source pin and report the expected 29 successes
and 23 refusals, but this run is not correctness evidence because the final
verifier rejected the corpus manifest before writing `smoke-verification.json`.

The fail-closed reason was a stale source label in
`current-corpus-manifest.json`: it still declared
`8702fd4db8723acceb7deb51bcb40ff66604bf10`, while the committed current
verifier expected d000. The observed traceback, command, session, and exit
code are recorded in `failure-terminal-31272.txt`.

The runner's EXIT cleanup removed the disposable Cargo target after the
verifier failure. `target-inventory.txt` records the exact target path and the
post-exit zero-file inventory. No target bytes are retained here.

`captured-source-16cf1e3c6.bundle` is the source bundle retained with the
successful repaired smoke. Its frozen ref is
`16cf1e3c612a14bd3517cd1b57dd57fce2fb6c35`; that commit descends through the
failed `901994438` checkout and the bundle therefore contains the failed
checkout's source objects. It requires d000 as its exact prerequisite. The
bundle SHA-256 is
`f3d273ad50647aa0daa7ea7dd809909c7c25a510fd344c535bd32f00b65e24a8`.

This rejected attempt must remain separate from the repaired smoke under
`../smoke-after-d000d977b/`. It contains no verification receipt and makes no
correctness, timing, or performance claim.
