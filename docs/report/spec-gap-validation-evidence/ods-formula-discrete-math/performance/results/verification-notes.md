# Retained verification notes

The final profile capture is retained under this directory with 330 baseline
rows and 1,530 candidate rows.  The frozen `verify.py` has one spelling typo
in its expected case tuple: it says `gstep`, while the harness emits the
canonical case suffix `gestep`.  [`verify_retained.py`](verify_retained.py)
SHA-checks the frozen verifier, loads it from its original `__file__` and
import path, changes only that tuple entry in memory, runs the original
verifier, and records the actual result in
[`verification-receipt.json`](verification-receipt.json).  No source,
harness, lock, contract, or profile input was changed after capture.

The matched DSUM control has a scoped RSS disposition.  Its evaluate median is
2,916 KiB at baseline and 3,120 KiB for the candidate (+204 KiB, +7.0%); its
parse-evaluate median is 2,924 KiB and 3,092 KiB (+168 KiB, +5.7%).  Raw RSS
distributions overlap but show the same median shift.  Allocation calls,
requested bytes, released bytes, peak live bytes, and result-live budget are
unchanged in both phases, so this is recorded as a process-RSS cost of the
feature batch with cause unresolved and no heap-growth evidence.
