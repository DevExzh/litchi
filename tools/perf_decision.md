# Performance decision utility

`perf_decision.py` derives one fail-closed qualification decision from a
strict facts document produced by an evidence auditor. The input records
separate statuses for completeness, custody, correctness, readiness,
performance, and benefit. It contains provenance references but no verdict;
the utility derives the verdict from those facts.

Run it with explicit paths:

```sh
python3 tools/perf_decision.py \
  --input /path/to/decision-input.json \
  --json-out /path/to/fresh-decision-output.json
```

The output path must be fresh. Existing files and symlinks are refused, and
the input is fully validated before any output is created. Missing evidence
produces `incomplete`; custody or readiness failures produce
`uninterpretable`; complete interpretable gate failures produce
`reject_candidate`; all passing gates with a declared benefit produce
`adopt`.

Exit status `0` means that valid facts were read and a decision was derived,
including `reject_candidate`, `incomplete`, and `uninterpretable`. Inspect the
JSON `decision` and `candidate_adoption` fields for the disposition. Exit
status `2` means invalid input, refused output, or another command-line error.

The module is an aggregation layer. Domain validators remain responsible for
raw packet parsing, correctness checks, thresholds, and evidence custody.
