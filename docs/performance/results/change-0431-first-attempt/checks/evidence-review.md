# 0431 evidence and replay review

This is an independent source-only review of `compare.py`, `verify.py`,
`verify-report.py`, `capture.py`, `check.py`, `replay.py`, the frozen
protocol, the build metadata, and the retained reports and receipts. I did
not run the comparison, verifier, portable replay, scripts, builds, tests, or
CPU workloads.

This sibling is the complete first-attempt bundle: both roles have sixteen
receipts in frozen order, the comparison is retained, `SHA256SUMS` is
present, and `checks/initial-portable.json` records a passing terminal
portable replay. It is historical evidence for the older capture behavior.
The main bundle points here while recapturing the after role with the refined
`after-v2` build; the first-attempt after result is complete and must not be
described as pending.

## Findings and disposition

### High: the evidence-driver hash chain was incomplete (resolved)

The original gap was that role receipts bound `capture.py` and
`verify-report.py`, while the analysis/replay drivers were outside the chain.
The current `replay.py` closes that gap for a sealed bundle: it requires
`SHA256SUMS`, records hashes for `replay.py`, `compare.py`, `verify.py`,
`check.py`, `capture.py`, `verify-report.py`, `seal.py`, and `protocol.json`
before replay, runs the portable check with its running log outside the
bundle, and checks those hashes and the inventory remain unchanged. The
current `verify.py` checks every command receipt's `driver_sha256` against
the retained `check.py`.

The terminal receipt deliberately uses `command` rather than `argv`, so it is
not misclassified as a `check.py` command receipt. The retained
`initial-portable.json` has status `pass`, zero exit code, all driver hashes,
and `drivers_unchanged` and `inventory_unchanged` both true. Its log and
receipt are covered by the sealed inventory. This resolves the source and
historical bundle finding; the final refined comparison belongs to the main
bundle.

### Medium: mandatory GNU-time RSS endpoint (resolved)

`capture.py` invokes `/usr/bin/time -v`, whose retained resource logs begin
the `Maximum resident set size (kbytes):` line with a tab. `compare.py` now
calls `line.lstrip()`, collects matching labels, and requires exactly one
matching line containing decimal digits before converting KiB to bytes
(`compare.py:229-239`). This accepts the retained raw-or-`.gz` endpoints and
rejects missing, duplicated, or malformed endpoint lines. The phase RSS scope
remains setup-inclusive endpoint observation; unavailable procfs points remain
unavailable rather than becoming a memory claim.

## Contracts that are sound

`distribution` uses linear-interpolated p50/p95/p99 values over each retained
30-sample process and a deterministic 2,000-resample, SHA-256-seeded median
bootstrap 95% interval for timing and rate metrics. Managed-resource and RSS
endpoints are summarized descriptively without bootstrap or causal
performance claims. The external GNU-time maximum RSS value is one
setup-inclusive endpoint per process; phase RSS values are lifecycle endpoint
observations.

Cross-role input matching is separated from output identity.
`input_identity` includes source and destination archive hashes and sizes,
the corpus manifest, and all producer gate values, while output hashes and
lengths remain in `output_identity`. Same-role R1/R2 output identities must
be stable, but before/after wrapper identities are intentionally allowed to
differ. Cross-role checks still require matched input and configuration, so a
new wrapper hash is not treated as an input mismatch.

Receipt custody is structurally sound for each retained report, log, and
resource file: every receipt names exactly those three artifacts and checks
their raw digest and byte count, including lossless `.gz` replay. The build
metadata binds the binary, protocol, source manifest, revision, and report
verifier. The source manifests compare every `tools/perf-baseline` Rust,
TOML, and lock entry between roles, preserving the timed harness boundary.

The source-bound semantic boundary is stated honestly in
`source-oracle-audit.md`: report verification checks the producer's retained
gate values and matched identities, but portable replay does not rerun Rust
readers against exported candidate archives. No output-wrapper equality or
independent archive revalidation is claimed, and no archive exporter is
required by this protocol. Decoded and physical untouched-member checks
remain producer-side evidence rather than portable replay results.

## Historical scope

The complete first-attempt comparison records the delayed-range API median
regression of about 8.8–10% and the bytes repeat-median trigger. Those are
retained observations, not a release or optimization claim. This attempt
used the older 16 KiB compressed-capture behavior; the main bundle is
recapturing the after role with the bounded 64 KiB refinement. The historical
comparison therefore remains valid evidence of that attempt while the main
bundle owns the refined result.

## Acceptance boundary

This sibling has complete historical captures and a passing portable replay.
The final refined comparison must come from the main bundle, with its own
after-v2 matrix, terminal replay, and final inventory. Editing this review
file changes a sealed artifact, so the sibling inventory/replay custody must
be refreshed before treating the edited tree as sealed.

## Root execution after source review

`precleanup-portable` passed the sealed role/comparison replay, isolated export,
and numeric/output/bound-verifier mutation rejection. Task scratch cleanup
then passed; the main bundle retains the cleanup inventory and command receipt.
The preceding pending language records the source review's earlier state.

`aftercleanup-portable` also passed after task binary/draft removal. The final
required-pass policy includes both terminal replays.
