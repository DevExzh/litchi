# Verifier negative reproductions

These historical notes preserve failures found while hardening the evidence
verifier. They are not final measurements or passing receipts.

- `/tmp/0454-fake-external-report.json` was rejected by the verifier.
- `/tmp/0454-forged-external-report.json` reached old assert-only checks and
  was accepted before the independent ZIP oracle was installed.
- `/tmp/0454-forged-range-report.json` reached old assert-only checks with
  `requested=2`, `returned=1`, and `short_reads=0` before monotonic and
  short-read validation was installed.
- `/tmp/0454-normal.out` and `/tmp/0454-optimized.out` recorded old verifier
  output; optimized `python3 -B -O` accepted `samples_raw=[]` because bare
  assertions were disabled.

The current verifier uses explicit `VerificationError` checks and the
independent stdlib ZIP oracle. These notes do not represent fixture or output
capture.
