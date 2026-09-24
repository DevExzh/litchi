# Staged-apply DOCX `stylesWithEffects` correctness smoke

This directory retains the fresh after-source correctness smoke for production
commit `506e6f8e5fb94a14c390fc262b152c773e8990b1`, captured from clean harness
HEAD `0aa99593cb609487cf85230c361be23ca34d4367`. The runner passed all 52
lanes. The receipt source identity, fixture/member hashes, typed refusals,
source readback, opaque preservation, inverse checks, and no-output checks are
verified by `smoke-verification.json`. This is a correctness prerequisite for
a future matched profile; it contains no timing matrix or speedup claim.

The external run produced 167 raw files totaling 1,765,619 bytes: 52 JSON
receipts, 52 `/usr/bin/time -v` sidecars, 52 empty stderr logs, and build,
source, metadata, command, and provenance receipts. The raw copy is recorded
in `retained-files.json`; every retained raw file was compared with the
external result byte-for-byte before cleanup. Root verified all 171 retained files against their committed Git bytes, then
removed the duplicate external raw directory. The clean source worktree remains
available; `REPLAY.md` describes restoration of the raw receipts. The disposable Cargo target was
removed only after the successful sentinel; its binary provenance remains in
`smoke-build-provenance.txt`.

The source bundle is `captured-source.bundle`, SHA-256
`a1a79dc80e5870593849f771fb81192b5f7e2b92845a5b3a486ffbbdfecb5d23`. It
contains the after smoke checkout and requires production commit
`506e6f8e5fb94a14c390fc262b152c773e8990b1`. The retained before profile and
the earlier 8702-source correctness smoke are separate evidence and were not
rewritten.

No matched before/after performance conclusion can be drawn from this smoke.
The future after profile must use the reviewed matrix, fixture bytes, lockfile,
resolved dependency closure, toolchain, flags, phase schema, and allocator/RSS
definitions from `after-capture-plan.md`.

`root-verification.json` is the independent root verifier result. Root restored
all 167 raw files solely from commit `421ff35ce` to the recorded external path,
ran the frozen `0aa99593` verifier, and confirmed byte equality with
`smoke-verification.json`. The temporary restoration was removed afterward.
