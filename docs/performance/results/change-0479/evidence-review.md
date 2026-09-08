# 0479 evidence review

Reviewed the frozen `analyze.py`, `verify.py`, `test_evidence.py`, and
`closure.py` against the Rust report schema, the capture/pilot protocol, and
the retained custody records. The final report schema and focused evidence
tests are aligned: the `analyze-final` receipt exits 0, the
`evidence-tests-final` receipt exits 0, and both report the source unchanged.

The concrete issues found in the earlier review are resolved in the current
files:

- `SINK_ID` and `SOURCE_ID` match the strings emitted by the Rust harness.
- Main-member decoded SHA-256 and CRC32 are checked against the independent
  XML oracle. The opaque member's deterministic decoded SHA-256 and CRC32 are
  also checked.
- Pilot reports are checked for binary/config identity, one-sample shape,
  sample invariants, and frozen-before-formal-capture chronology.
- Formal capture argv now binds the executable path to the recorded binary.
- Total and phase allocator retention rules, RSS scope, and the repeat-1
  growth basis are explicitly represented.
- The required final labels are `evidence-tests-final` and `analyze-final`,
  and the ledger/verifier agree on that set.

Two closure actions remain before the bundle can be sealed:

1. Remove `docs/performance/results/change-0479/__pycache__/` and its `.pyc`
   file. `seal.py` and `portable.py` explicitly reject any `__pycache__`
   path, so sealing will fail while it remains.
2. Bind and verify the generated `profile-summary.json` from
   `profile-analysis.py` in the final closure record. `closure.py` currently
   validates the raw profile inputs and the three profile reports, but does
   not yet compare or hash-check `profile-summary.json` against an independent
   `profile-analysis.py` derivation. The profile analysis parser itself has
   the needed reconciliation checks; the custody link is the pending part.

`evidence-validation.json` is also still required by the complete closure
path. It is an execution/receipt artifact rather than an analyzer defect and
must be generated after the final helper set is frozen.

The retained whole-archive SHA-256 remains a producer/runtime oracle because
the archive bytes are not retained in the report. Independent XML, semantic,
member-metadata, order, sink, and inverse checks cover the declared evidence
boundary; this limitation must remain explicit in the final report.

## Coordinator closure

The owned Python cache was removed. `closure.py` now recomputes the profile
summary and checks compressed/raw perf identity, all profile reports, final
live/cleanup custody and the helper/test-receipt binding. All 18 Python helpers
are bound in `evidence-validation.json`. The complete verifier passes the
24-report/720-sample bundle and all 16 required final gates. After authenticated
runtime cleanup, a copied bundle passes and eight independently resealed
corruptions are rejected. `portable.json` retains their exact diagnostics.
