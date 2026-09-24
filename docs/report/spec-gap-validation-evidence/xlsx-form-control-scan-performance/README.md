# XLSX form-control namespace/event scan fusion

This directory is the final replay bundle for the bounded `select_mce`
optimization.  It contains one final source delta, the exact source and
harness locks, the final fixture bytes, and one final baseline/candidate
receipt pair.  Earlier candidate revisions are intentionally absent.

The pinned source base is commit
`1ae5d7b4504a248ea24ac52cb2bdd1f08859ddd7`.  The reviewed production commit
is `ab5707afc71c1e9e3f724459a482867f50aa8163`; its `owner.rs` blob is the same
as the candidate reconstructed by this bundle.  Replay uses the base commit
plus [`source-delta.patch`](source-delta.patch), so it does not accidentally
mix unrelated changes from a later worktree.

The candidate source identities are:

| source | git blob | SHA-256 |
| --- | --- | --- |
| baseline `owner.rs` | `5ee687c00133761fddc6ee3cd4be3cc9ebea8e67` | `ac8f6dfadd81d1a584488d0dcb3fcb3ef6ceaba75d51f99dc9fab0843df0bb3a` |
| candidate `owner.rs` | `541c1a3ca4d4b742ac6a134aae8dc942f61404ed` | `778c510def4174e196ae277ec92127337f49bb28b8834f9bbcdb4b889b60d5c4` |

The patch SHA-256 is
`744139737b5817a2393535abe95b2842b4db95b02aebcc2ff06f50f2ecd34bbb`.
The source lock SHA-256 is
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`; the
harness lock SHA-256 is
`d5814a10e3a006ab2457c10fda9af3d9cd33c1245c744ab649711acab5b0ebd9`.

## Replay

Run from a repository that contains the pinned base commit:

```sh
TMPDIR=/var/tmp \
  ./docs/report/spec-gap-validation-evidence/xlsx-form-control-scan-performance/replay.sh \
  "$PWD" \
  /var/tmp/form-owner-scan-replay-final
```

The script restores two clean source trees with `git archive`, verifies the
baseline and candidate owner blobs, installs the retained source lock, builds
the same harness twice with separate targets, and runs the same five fixture
cases for seven repetitions.  It writes baseline and candidate replay
receipts, build logs, binary hashes, and `replay-metadata.json` to the output
directory.  The metadata checks that correctness records and all source
readset counters match; timing and allocator fields remain separate evidence.

The final retained historical receipts are:

- [`receipts/baseline-final.jsonl`](receipts/baseline-final.jsonl), SHA-256
  `49edbd21cbc7a5b677c8e5e34f0144a0941a5894481eb4dcf59108af4783005f`;
- [`receipts/candidate-final.jsonl`](receipts/candidate-final.jsonl), SHA-256
  `84e6c811fec3ec37bcb3c714afc983725d3834511672c54150844b88a3ceaf6d`.

A clean source reconstruction/readset replay completed with five correctness records and 294
receipt records for each source.  Its output metadata records the exact
receipt and source hashes.  The four admitted cases and one refused case use
only the fixture bytes retained under [`fixtures/`](fixtures/); no private
target directory or earlier receipt is required.

The production commit `ab5707afc` gates are retained in
[`xlsx-form-control-scan-final`](../xlsx-form-control-scan-final/):

- `cargo test --locked -p litchi-xlsx --lib --offline`: 1,171 passed;
- `cargo test --locked -p litchi-xlsx --test form_control_owner_read --offline`:
  39 passed;
- library and integration clippy with `-D warnings`: passed;
- rustdoc with `-D warnings`: passed;
- rustfmt and `git diff --check`: passed.

The matched allocator harness observed a deterministic reduction on the
16-control source projection from 57,796 to 57,788 allocation calls and from
34,526,604 to 34,520,703 requested bytes.  Final paired process medians were
workload-scoped observations only; this bundle makes no suite-wide or
production speedup claim.

The replay runs the allocator harness, not the Cargo test suites. It uses fresh archived source trees and separate targets on the same host; it is not an independent timing study. Release elapsed/RSS observations include global allocator instrumentation, sequential fixture order, and ordinary cache/scheduling effects, and are excluded from replay identity.

Root independently reran the clean reconstruction; [metadata](root-replay/replay-metadata.json) and [allocation comparison](root-allocation-verification.json) retain the results. All correctness, source-read, live-byte, and peak-byte fields match. The 56 admitted projection records save 8 allocations, 8 deallocations, 25 reallocations, and 5,901 requested bytes each. The 14 refusal records in this replay save 8 allocations, 8 deallocations, 26 reallocations, and 6,117 bytes; their totals differ from historical receipts. These are measured per-lane deltas, not a guarantee of deterministic allocation totals across runs.

Independent bundle review verified source/replay integrity and required the committed test-count and replay-methodology qualifications above.
