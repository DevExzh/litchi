# Current-source DOCX stylesWithEffects correctness gate

The fresh correctness smoke passed on frozen harness commit
`637082e318ae6bee2bd47ef317e4a417c1176996`, with production source pinned to
`8702fd4db8723acceb7deb51bcb40ff66604bf10`. This is a correctness prerequisite,
not a timing capture or a performance improvement claim.

The run covers four fixtures and 52 separate-process lanes with one sample
and no warmup: 29 expected successes and 23 expected refusals. The verifier
checks native fixture/member hashes, exact-fit and one-under controls, receipt
source identity, and phase accounting. All 52 process sidecars report exit 0
and their stderr files are empty. Before/after source manifests, Cargo
metadata, and binary hashes agree. The source closure contains 790 files and
21 extra inputs.

`smoke-verification.json` is the original runner result;
`root-verification.json` is an independent root verifier replay. Independent
review also approved the complete receipt set. `retained-files.json` records
the length and SHA-256 of each of the 167 original files copied here, totaling
1,767,604 bytes. Raw logs are retained unchanged, including whitespace and
original external paths.

The build has three nonfatal dead-code warnings for profile-only allocator
helpers (`phase_peak`, `phase_delta`, and `begin_phase_window`) compiled into
the smoke harness. It has no build errors. This run is not a strict-Clippy
gate.

The preceding `133c4ede0` attempt is not approved: its 52 receipts incorrectly
reported historical source `d1f299d00` and the verifier refused them. The
repaired harness emits the current source pin; this directory contains an
entirely fresh run, not rewritten receipts from that attempt.

Root launched the frozen `run_current_smoke.sh` with
`PYTHONDONTWRITEBYTECODE=1`, results at
`/var/tmp/litchi-docx-styles-effects-current-smoke-results-637082e31`, and a
fresh target at
`/var/tmp/litchi-docx-styles-effects-current-smoke-target-637082e31`.
The successful runner removed its owned target. The frozen checkout remains
separate from the shared feature worktree. No timing matrix ran.

The 23,007-byte `captured-source-637082e31.bundle` preserves the frozen source
history, requiring main-history commit `8702fd4db` as its prerequisite. Its
SHA-256 is `13ec534e1aa451c62b43ce8cc5e408e62af20bcf5cac8d4adc371e36c63e6b8c`.
Root verified the bundle against this repository. See [REPLAY.md](REPLAY.md)
for restoration commands and the distinction between frozen source hashes
and subsequent explanatory documentation changes.
