# ODS text-functions performance evidence

This directory owns the reproducible profile for the 26 OpenFormula 1.4
section 6.20 text functions. It compares baseline
`8f09231e36982248eface4d599432143a67f6e49` with the frozen candidate using
the retained gate lock
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The reviewed contract is pinned at SHA256
`b03ec074efe30c4261509a70b6188010d74a9d48d8bc92a57ad39a2db4a3b611`.

The [plan](PLAN.md) defines 28 matched controls and 133 bounded candidate
cases; the generated [case matrix](case-matrix.json) records every name and
read bound. It includes borrowed and owned concatenation growth, scalar and large
Unicode inputs, 64-cell borrowed references, mapped matrix outputs, typed
refusal, sticky cancellation, resource limits, worst-case searches, REPT
growth, and ASC/JIS width conversion. Text lanes report input/output bytes and
normalized byte throughput in addition to elapsed time, work, resolver reads,
allocation, live-memory, and RSS measurements.

The six-position fraction workload uses a case-local scalar `max_steps=2,000,000`
caller ceiling while exercising the bounded continued-fraction kernel for a
999,999-denominator maximum. All other cases use the evaluator's default scalar
limit; the existing small-budget refusal coverage remains unchanged.

The runner refuses capture until the text contract, oracle/generator/goldens,
native provenance, retained lock, and candidate freeze are present and hashed.
It preflights all named cases before timing and uses three warmups and 15
fresh child processes in both evaluator phases. Raw stdout, `/usr/bin/time -v`
logs, source/profile manifests, and cleanup receipts remain under `results/`.
The package owns this README and plan, the runner, summarizer, verifier, and
harness. Production implementation, contract, oracle, and native evidence
remain owned by their respective agents.
