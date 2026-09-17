# Change 0675 integration evidence

This packet closes the queue authorized by change 0652 and tracks the resumed
worktrees. It makes no independent performance claim.

- `run-integration.py`: reproducible merged-branch gate commands; run from the
  repository root. `LITCHI_GATE_OUTPUT` selects the evidence directory and
  `LITCHI_GATE_ONLY` optionally selects comma-separated gates for a documented
  retry. No selector means all gates.
- `merge-log.txt`: source commits and cherry-picked commits.
- `cleanup.json`: clean completed-worktree removal results.
- `integration/`: merged-branch command output and exit statuses.
- `validation-summary.json`: final test totals and all-gates-passed result.
- `integration-attempt1/`: the retained XLS inverse failure before the padding fix.
- `pre-final/` and `pre-cfb/`: intermediate merged check/Clippy/test results,
  including the Clippy finding and its correction before final integration.
- `documentation/` and `claims-followup/`: strict/structural claims, coverage,
  boundaries and claim-checker tests.
- `review-notes.md`: findings that required implementation follow-ups.
- `preserved-historical-logs.json`: existing logs exposed by the general ignore
  exemption; these were neither staged nor deleted.
- `artifact-cleanup.json`: removed build outputs owned by completed agents.
- `docx-mce-regression.txt`: the complete 40-test tail-append suite after
  replacing the obsolete namespace-amplification expectation.

Each implementing record owns its own measurement evidence. This packet's
validation does not promote any measured observation to a broader performance
claim.

Captured terminal logs normalize trailing whitespace and final blank lines;
commands, diagnostics, test outcomes and exit statuses are unchanged.
