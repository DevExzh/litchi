# Owned-value performance harness review

This is a source and receipt review of the candidate-only
`Evaluated::to_owned` harness. No Cargo command or profile run was performed
for this review.

Reviewed harness inputs:

| file | SHA-256 |
| --- | --- |
| [`Cargo.toml`](Cargo.toml) | `491af76b4343589f96d5171b551c564382953774ebaad63c31c7dda297d20aa7` |
| [`Cargo.lock`](Cargo.lock) | `5d7517a80576661f72a1930cf19b7d6c20e2bf01db61ef609a73eea5b7307dbf` |
| [`run.py`](run.py) | `7b03069fa2e9b93ecba6f7660c2f70adf116aef8a2f55a4f6ea1bc382b31b636` |
| [`src/main.rs`](src/main.rs) | `c6173389a3da6ff3b3d33b5879c042bcd5f673e3b954a9313908200642515e2b` |
| [`README.md`](README.md) | `5d1569a419735ca732ed99e1a8e43b4a08f2454f9d49badadab95b5ec6bf7bd1` |

## Confirmed measurement boundaries

* `own` prepares parsing and evaluation before the timer, then times only
  `to_owned`, owned-view checksum traversal, and result drop. The separate
  `setup` phase measures repeated parse/evaluate preparation. The
  `parse-evaluate-own` phase includes all three operations and is correctly
  documented as a broader end-to-end control rather than a comparable
  replacement for `own`.
* `elapsed_ns` is divided by the configured repeat count. Allocator calls,
  requested/released bytes, work, retained budget usage, and live-byte peaks
  are deliberately retained as batch totals or maxima. `owned_reserved_bytes`
  is sampled while the result is live; `peak_live_delta` is a raw high-water
  delta and is not divided by repeat. GNU `time -v` RSS is process-wide. These
  meanings are consistent in the runner and README.
* The timer starts after per-case source construction, limits, and the one
  preparation result are established. In `own`, the preparation result lives
  on a separate execution budget, so the timed execution budget measures the
  owned result reservation. End-to-end measurements intentionally include the
  per-iteration evaluated result and owned result together. Successful and
  refusal paths drop owned results before the next repeat.
* The counting allocator resets counters before each sample and resets the
  live-byte peak to the pre-timer baseline. The wrapper accounts allocator
  requested/released layout bytes and a logical live-byte high water mark;
  this is not physical allocator overhead, RSS, or a general zero-copy claim.
  In particular, a `realloc` is represented by its net live-byte change, so a
  possible old-plus-new physical overlap is outside this metric.
* The four refusal cases have the intended phase policy. Setup uses the
  normal preparation limits and succeeds; `unicode-limit` and
  `array-limit-memory` refuse with `resource-memory`, `array-limit-work` with
  `resource-work`, and `array-cancelled` with `cancelled` in the copy phases.
  The runner checks status, typed failure, and p50 success/refusal counts and
  rejects partial or mixed phase output.

## Current corpus blocker

The supplied preflight capture contains six real failed lanes: both
`duplicate-list-4096` and `three-d-list-4096` fail in `setup`, `own`, and
`parse-evaluate-own`. Each child prints its config and then exits before a
result with:

```text
preparation evaluation failed: Work budget exceeded in ods-formula-evaluation: observed 1000001, limit 1000000
```

These cases are marked `expected_success=true`, so the runner correctly rejects
them rather than classifying the evaluator refusal as an intended harness
case. The failure is in the current production reference-list preparation
work (quadratic union accounting), not in the copy-phase refusal map. The
4,096-reference cases should remain success controls and be rerun only after
the evaluator-side bounded-work fix; increasing the harness limit would hide
the production boundary under review.

## Evidence limitations and follow-up

* `checksum_owned` traverses every array cell and reference area, but for a
  reference it includes only owner presence, area bounds, and area count. It
  does not include source IRI, sheet names, subtable names, endpoint kinds,
  or absolute/quoted lexical markers. It is therefore a useful consumption
  guard for these fixed cases, not an independent proof of complete deep-copy
  semantics. The owned-value integration tests remain the semantic authority
  for those fields.
* The checksum is not compared with independently generated expected values;
  the runner validates only repeat success/refusal and the child-reported
  checksum fields. A future harness verifier should bind expected scalar,
  array, text, and reference metadata values outside the timed region.
* `child_validation` checks required fields, nonnegative numeric values, typed
  failure, and p50 counts, but it trusts child-reported `input_bytes`, shape,
  element count, mode, and value kind. The source and harness hashes make the
  current fixed corpus reproducible, while an independent metadata table would
  make the runner itself stronger.
* The runner hashes the binary and all five harness inputs before and after
  the serial run, but the supplied source binding is a one-time manifest hash.
  If concurrent source-manifest mutation must be ruled out, record and compare
  that manifest hash again at completion.
* The normal runner defaults are bounded (three warmups, fifteen samples, and
  fixed per-case repeats), but the child accepts arbitrary warmup/iteration
  values and direct callers can provide arbitrary repeat values. Evidence
  commands should retain the fixed defaults or add explicit finite CLI caps.
  With fifteen samples, p95 and p99 are coarse order statistics and should not
  be presented as robust population-tail estimates.

The harness is suitable as candidate-only diagnostic evidence once the two
4,096-reference preparation lanes are made runnable. It supports memory and
ownership accounting, but it cannot substantiate a baseline-versus-candidate
speedup claim because the public owned conversion is absent from the baseline
API and the checksum is intentionally not a full semantic oracle.


## Root follow-up after this source review

The root added source-manifest retention and an after-capture hash fence, plus
runner iteration caps. `runner-checks.json` records successful checks against
actual GNU time whitespace and a mocked child that mutates the source manifest.
The main executable now validates independently expected fixture content and
reference lexical metadata before timing, includes those fields in checksums,
and caps repeat counts. A root debug build passed; the third preflight passed
48 lanes and retained the same six large-list preparation failures. This follow-up
records executed integration evidence, not a second independent source review.
README documentation was expanded after that capture; the executable and runner
sources were not changed by that documentation update.
