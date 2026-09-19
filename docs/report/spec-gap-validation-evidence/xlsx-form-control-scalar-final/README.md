# Existing XLSX form-control scalar edits

This batch adds selector-first scalar edits for an existing worksheet control,
with paired control-properties and VML publication. Ordinary workbook edits and
the source-backed editor support exact no-ops, reversible patches, owner/source
checks, and atomic refusal. Ordinary edits expose structured same-worksheet
conflicts and a logical `PackageChange::FormControl` record.

The bounded profile admits one control per worksheet batch. Creation, removal,
list edits, graph-changing formulas, execution, and rendering remain outside
this batch. Edited files have library save/reopen evidence, not native Office
acceptance evidence.

- [Contract and resource ownership](contract.md)
- [Independent review](review.md)
- [OPC boundary review](opc-review.md)
- [Scalar mirror evidence](../xlsx-form-control-mirror-evidence.md)
- [Performance workflow and measurements](performance/README.md)

The final gates run in an isolated checkout containing the selected batch over
`967c04c1f`, excluding unrelated pending workspace changes. `gates/run.py`
records the compiler, lockfile, source hashes, command output, and exit codes.
Display-only trailing spaces in command logs are stripped before hashing.

The baseline drawing test also needs LibreOffice's
`sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx` under
`3rdparty/libreoffice-core`; its input hash is included in the gate manifest.
Full-crate formatting reports existing differences in the unchanged
`drawing_svg_read.rs`. The runner verifies that file is identical to the base
commit and that no other file appears in the format failure, then separately
requires every selected batch file to pass `rustfmt --check`.

## Validation

The isolated XLSX suite passed 1,831 tests across 74 result groups, including
39 owner-read and 32 scalar-lifecycle regressions. The OPC suite passed 619
tests across 22 result groups. Its existing independent ZIP64 corpus test is
ignored without `LITCHI_0415_PYTHON_ZIP`; this batch does not claim that case ran.
See [gate results](gates/results.json), [verification](gates/verification.json),
and [frozen batch hashes](gates/freeze.json) for the exact checked source.

All-target Clippy and rustdoc passed with warnings denied. Crate-boundary,
selected-file formatting, and diff-whitespace checks passed. Before/after source
manifests are identical, and the retained log hashes were independently
recomputed. The full-format baseline exception is recorded separately, so
`all_commands_passed` is false while `all_required_checks_passed` is true.

## Performance evidence

The final capture contains 360 samples across three retained fixtures and eight
lanes, with three warmups and 15 samples per fixture/lane. All no-op, changed
reopen, and inverse correctness checks passed. Source-backed scalar
commit/publication/reopen medians were 3.59–13.75 ms on this machine and corpus;
these are workflow observations, not a before/after optimization claim.
[Raw data and percentile report](performance/results/final-capture-20260919/report.md)
include allocation, retained live-byte, RSS, and `ReadAt` measurements.

Run `python3 verify.py` from this directory (or by absolute path) to independently
recompute the profile statistics and verify source/receipt identity. Final
verification passed after capture; no production source changed afterward.
