# 0837 — rejected fresh compression; retained ZIP retry fix

The fresh-only OPC compression candidate was **rejected at correctness
qualification**. The retained production change lets the bounded owned ZIP
entry adapter propagate retryable `Interrupted` writes without poisoning the
entry. No performance improvement is claimed. See the
[change report](../../0837-zip-interrupted-write-and-fresh-compression.md).

## What was executed

- Baseline OPC tests, baseline probe build and six two-sample preflight cases.
- Candidate OPC short/Interrupted stream test: failed, then passed after the
  bounded ZIP write adapter was corrected.
- Candidate formatting, compilation, Clippy and rustdoc: passed. The dependent
  test suite stopped at a durable DOCX inverse exact-archive failure.
- The OPC candidate was reverted. The first `final-*` gate sequence reused
  candidate artifacts because restoring with `copy2` retained old source mtimes.
  That sequence is invalid as final verification and remains preserved.
- The corrected `final-v2-*` sequence rebuilds affected source after refreshing
  mtimes. Its immutable source inventory, receipts and full test results define
  final verification. `red-green.json` binds the direct old-adapter failure and
  corrected Store/Deflate regression run.

The final affected-owner suites pass 6,847 tests with zero failures and 38
ignored. Closure validates 37 recorded command receipts. Cleanup removed
20,289 files / 6,275,196,049 logical bytes from the two owned roots; offline
closure and diagnostic replay pass afterward.

Only six baseline preflight reports / twelve measured operations exist. The
planned candidate release build, qualification matrix, admission and native
comparative captures **did not run**. The plan and unexecuted comparative tools
(`capture.py`, `admit.py`, `analyze.py`, `decide.py`, `audit.py`) remain as trial
artifacts, not evidence of a passed comparative gate. Initial script and probe
versions, build failure, invalid patch drafts and the mislabeled baseline
`candidate-focused` receipt also remain inspectable.

## Core evidence

- `rejection.json`, `rejected-source.json`, `candidate-rejected.patch` and
  `sources/candidate/`: exact rejected implementation and reason.
- `review.md`: both independent read-only reviews and the correctness finding
  that supersedes the provisional patch assessment.
- `inverse-failure-analysis.json`, `diagnostics/`: assertion-derived ZIPs show
  equal logical bytes but a different compressed `word/document.xml`.
- `final-v2-quality.json`, `final-v2-quality-source.json`, `sources/final/`:
  source and affected-owner final gates; no OPC compression change remains.
- `red-green.json`: direct baseline failure and final pass for the retry fix.
- `closure.json`: command census, test counts, source deltas and baseline probe
  custody; `cleanup.json`: removal of the two exclusively owned temporary roots.
- `seal.json`: exact owned files and hashes, excluding the seal itself. The
  commit's owned path set must equal these paths plus the seal.

## Reproduction and limits

From this revision, reproduce the retained production regression with:

```sh
cargo test --offline --locked -p soapberry-zip --test streaming_interrupted_write
python3 -B docs/performance/results/change-0837/diagnose.py
```

`diagnose.py` uses relative retained artifacts and Python's independent ZIP
reader. It proves the physical discrepancy without rebuilding the rejected
candidate. `python3 -B docs/performance/results/change-0837/close.py --check`
replays custody in the recorded workspace at the final source state, including
its unrelated-file preservation checks. Deleted executable hashes resolve
through the cleanup manifest. Historical capture scripts intentionally require
the original base revision and exclusively owned roots; they are not general
post-commit replay commands.

The baseline release uses the same pinned workspace dependency versions as its
probe lock. Its original failed build is distinct from `baseline-v2`, the
successful binding used by `build-baseline.json`. The probe timer is prepared
OPC `to_bytes`; it is not a complete DOCX/XLSX/PPTX create/edit/save lifecycle.
No candidate time, cold-cache, filesystem, concurrent, allocation or Office
interoperability result follows. iWork and unrelated workspace changes are
excluded. The broader performance goal remains active.
