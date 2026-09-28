# Quality scope

The root coordinator runs `quality.py before`, `quality.py probes`, and
`quality.py after` serially, retaining each attempt's commands, environment,
source hashes, logs, exit codes and timestamps. The before and after legs run
formatting, all-feature/all-target checking, all-feature tests and doctests,
warning-denied Clippy, warning-denied rustdoc, and the workspace dependency and
crate-boundary audit. This is the applicable PPTX crate quality matrix, not a
claim to rerun every format's tests or workspace-wide compilation.

The probe leg checks both independent manifests, in default and all-feature
configurations. It runs formatting, checking, tests, Clippy and documentation.
The optional allocator and capture-profile configurations are confined to
evidence binaries. All Cargo calls are offline and locked; jobs are limited to
two, incremental compilation is disabled, and development debug info is
disabled. The quality target is shared between sequential legs; each receipt
binds the actual source bytes so Cargo reuse cannot stand in for source custody.

The production differential oracle compares the candidate scanner with the
retained old scanner loop. Existing public Scene, MCE, malformed input,
preservation, transaction, patch, and package tests remain in the all-feature
suite. Native comparison is withheld until both qualification lanes accept all
19 workflows, and final adoption also requires the performance/resource policy.

The 0822 HEAD seal was replayed before this work. Its profile supplies the
hotspot evidence, while every before/after performance sample in this trial is
new. No historical timing observations are pooled into this trial.
