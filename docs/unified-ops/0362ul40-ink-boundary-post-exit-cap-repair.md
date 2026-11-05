# 0362ul40 — DOCX durable Ink post-exit cap repair

Date: 2026-10-24. This correction closes the aggregate resource-admission
race in the bounded Ink launcher. It does not rebuild, rerun, relabel or
qualify the accepted f318 DOCX artifacts.

## Finding and source binding

The prior launcher checked aggregate output and TMP bytes only while the
systemd-run process was live. A fast child could write beyond a cap and exit
between the last poll and process.poll(), leaving a complete receipt with no
violation. The fix is in
tools/ink_safe_runner.py at commit
1bc982f4495a16c30ac4ca2779351aec1e420781a6d6cb7c2b1f33ccff88fede
(17,630 bytes). It performs a final output/TMP snapshot after process exit,
uses the same strict > policy, preserves the return code and receipt
artifacts, and invokes the named .scope cleanup when that final snapshot
fails.

The updated test source is
tools/test_ink_safe_runner.py, commit
ee6ac475191373c7fe581c2767da55c98449db964b60a5ea5fc9e49e60095499
(10,397 bytes). The external source-preparation receipt was rebound to
21,715 bytes, SHA-256
68259481ceb41a58cc2371457f06a942f86cd62c2ba5f3ade6d5180a15a8620f.

## Source-only validation

    PYTHONDONTWRITEBYTECODE=1 python3 -B -m unittest tools.test_ink_safe_runner -v
    PYTHONDONTWRITEBYTECODE=1 python3 -B -m py_compile \
      tools/ink_safe_runner.py tools/test_ink_safe_runner.py

All 11 tests passed. The new cases are:

- a fast-exit child exceeding the output cap;
- a fast-exit child exceeding the TMP cap;
- output and TMP trees exactly equal to their caps.

No Cargo, native Ink binary, profiler, timing run, or fixture was invoked.
The accepted systemd-259 /bin/true and escaped-PGID cleanup preflights from
the preceding packet remain unchanged and external.
