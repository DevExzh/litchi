# ODS selector performance evidence

[`report.md`](report.md) compares an absolute baseline from commit `d5b5afdb4032b718c0f94c113a3b1b09f17ae6ba` with the candidate `index.rs` binary-search patch. The baseline and candidate directories retain the exact harness outputs, `/usr/bin/time -v` RSS records, host snapshots, source hash manifests, and replay command ledgers. [`comparison.csv`](comparison.csv) is the machine-readable paired result.

The paired run covers `parse` 64×64, `lookup` 32×32/64×64/256×256, `stage-one` 64×64, `stage-batch` 32×32/64×64, and `edit-batch` 32×32/64×64. Both candidates use the corrected valid-address fixture, one sheet, three warmups, and fifteen measured iterations with the process pinned to CPU 2. Budget `Memory`/`Work` counters are reported separately from operating-system maximum RSS.

The candidate source is reconstructible without a source archive: apply with `git apply --unidiff-zero` [`candidate/index.diff`](candidate/index.diff) to the pinned commit and verify [`candidate/index.diff.sha256`](candidate/index.diff.sha256), then check [`candidate/source-hashes.sha256`](candidate/source-hashes.sha256).

[Review and validation](review-and-gates.md) explains the search invariants and
preservation boundaries. [Root verification](root-verification.json) records
711 passing tests across 43 targets, strict Clippy, rustdoc, scoped formatting,
and independent checks of the paired raw measurements and source hashes.
