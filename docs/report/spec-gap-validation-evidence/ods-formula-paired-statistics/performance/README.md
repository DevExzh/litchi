# ODS paired-statistics performance evidence

This directory owns the reproducible profile for `CORREL`, `COVAR`,
`PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and `FORECAST`. The matched
baseline is `aa48eee68cb6ee0904523394d4ff8015dce6e595`; dependency resolution
uses the retained order-statistics gate lock
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.

The reviewed contract hash is
`9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd`, with
`INTERCEPT` using the included-constant regression. The [plan](PLAN.md)
defines 39 matched controls, including normal-reference and sensitive/extreme
inline regression controls for `AVEDEV`, `DEVSQ`, `KURT`, `SKEW`, and `SKEWP`,
and 66 bounded candidate groups.

The runner refuses capture if the reviewed contract, numeric oracle/goldens,
native provenance, retained lock, or candidate freeze is missing or changed.
It preflights every candidate case before timing either side, then uses three
warmups and 15 fresh children in both evaluator phases. Raw child JSON,
`/usr/bin/time -v` logs, allocator/work/read evidence, source closures,
profile-input hashes, and cleanup receipts remain under `results/` after a
capture. No save, recalculation, or cross-platform timing claim is made.

Owned paths are this README and PLAN, `run_profile.py`, `summarize.py`,
`verify.py`, `harness/`, and future `results/` or `attempts/` receipts. The
production implementation, contract, and native producer remain owned by
their respective agents.
