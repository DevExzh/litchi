# DOCX stylesWithEffects projection reuse correctness smoke

This directory retains a fresh bounded correctness smoke for the committed
projection-reuse candidate. The production source pin is
`d000d977b99e03f8542c7dae74acf767a91b1feb`; the frozen harness and pin-label
checkout is `16cf1e3c612a14bd3517cd1b57dd57fce2fb6c35`.

The smoke used four committed native fixtures and 52 separate-process lanes,
with one sample and no warmup per lane. It produced 29 expected successes and
23 typed expected refusals. The runner verification passed all 52 receipts:
semantic and opaque checks passed for every lane, expected success/refusal
outcomes matched, and all receipts carry the d000 source pin. The source
manifest has 790 package files and 21 explicit evidence inputs; before/after
manifests and Cargo metadata match. The verification receipt is
`smoke-verification.json` and the per-file hashes and sizes are in
`retained-files.json`.

The 167 raw files copied from the external result directory total 1,771,665
bytes. This includes every lane JSON receipt, `/usr/bin/time -v` sidecar, empty
stderr sidecar, build log, binary hashes, Cargo metadata, source manifests,
commands, provenance, and the runner verification. Raw files were copied
byte-for-byte and must remain unchanged. The binary digest recorded by the
runner is
`e2da314c96bddb85d0d47a4de290d372a0acaef352b21aaca29f743a3e72304e`.

The source bundle `captured-source-16cf1e3c6.bundle` contains the frozen
checkout history through `16cf1e3c6` and requires d000 as its exact Git
prerequisite. Its SHA-256 is
`f3d273ad50647aa0daa7ea7dd809909c7c25a510fd344c535bd32f00b65e24a8`.
`REPLAY.md` gives the bundle verification and checkout commands.

The successful runner removed its owned Cargo target after verification; the
absolute binary and target paths in the receipts are provenance paths, not
promises that the disposable target still exists. The external source result
directory was
`/var/tmp/litchi-docx-styles-effects-projection-reuse-smoke-results-16cf1e3c6-20260912-b`.

An earlier attempt at
`/var/tmp/litchi-docx-styles-effects-projection-reuse-smoke-results-901994438-20260912-a`
is retained separately as failed evidence: its verifier correctly refused the
old corpus source pin before producing a verification receipt. It must not be
combined with this successful run.

Root verified all 167 raw files against their external originals and verified
the source bundle. A separate clean checkout at the exact captured HEAD
replayed the frozen verifier successfully; `root-verification.json` is byte
identical to the runner's verification receipt. `root-replay-context.json`
records that checkout and its removal.

This is correctness evidence only. No timed profile samples, optimization
speedup claim, or native-application performance claim is made here. A later
profile prerequisite review remains a separate gate.
