# 0489 OPC candidate XML audit reuse evidence

This bundle evaluates reusing the successful initial candidate XML audit
inside one immutable prepared OPC splice plan. The [hypothesis](hypothesis.md)
records the proof and ADR obligations before source changes. The
[methods](methods.md) describe the matched 18-arm matrix, two binary roles,
two process repeats and 4,320 formal measured samples across both phases.

The before binaries are the retained 0487 executables. New build output uses
an isolated Cargo directory that is removed at batch completion. Benchmark and
fuzz executables bound by evidence are copied separately and retained.

The standalone source audit, initial candidate audit, per-pass byte/EOF/hash
verification, source freshness, resource accounting and final DOCX reopen remain
required. The full non-iWork performance goal remains open.


[The results review](results-review.md) accepts this scoped change with about
20–21% lower source-heavy owned/file medians and lower operation heap. It
explicitly retains small file-store tail regressions and four adverse RSS
observations. All candidate archive bytes match. The [cleanup](cleanup.json)
and [seal verification](seal-verification.json) record post-validation custody.
