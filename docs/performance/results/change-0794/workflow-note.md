# 0794 execution note

The first application check used `git apply --check --whitespace=error -p0`
after the archived patch had been finalized with `a/` and `b/` path prefixes.
The check failed with missing `b/crates/...` paths and made no production edit.
The surrounding shell did not fail closed between that Python check and the
following quality command, so the latter started against the baseline source.
Its `quality-0/source.json` is the authoritative source witness. That live run
was retained rather than interrupted or mislabeled as candidate validation.
The candidate application uses the normal strip-one prefix and will be
verified in a separate completed tool call before its quality command starts.
The public before/after measurement plan and policy are unchanged.

A separate helper count/layout diagnostic was declared and built after the
public baseline qualification but before any candidate application or timing.
Its frozen driver/source inventory is stored per leg. It measures no latency,
is outside the public adoption policy, and uses the same pinned quick-xml
0.41.0. Two reports each contain 42 measured helper iterations (14 counts × 3).
The public plus profile totals remain 259 reports / 5,599 samples; the helper
supplement adds two reports / 84 iterations in a different schema.

The shared source scope also motivated fresh DOCX/XLSX controls using the
existing harness. Their separate plan, commands, and rejection rule were
frozen before their baseline compilation and before candidate application.
This is an additional regression veto; no observed result was used to relax
the original PPTX adoption policy.

Candidate validation also retained three pre-measurement failures. Attempt1
caught unintended OLE common API visibility changes; only the original `pub`
visibility was restored. Attempt2 then exposed constructor shadowing in two
new test-local bindings per copy; these were renamed. Attempt3 rejected the
resulting long assertion lines under rustfmt; formatting was applied. Exact
failed candidates, source censuses, and logs remain archived. None of these
attempts produced candidate performance measurements. The final quality run
uses the corrected five-file candidate, with its manifest, patch, and application
receipt regenerated before measurement.
