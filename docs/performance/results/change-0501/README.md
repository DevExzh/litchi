# 0501 measurement bundle

## Completed result

The private payload hash reduction passes all 13 recorded crate/tooling gates.
The matched matrix contains 16 passing children and 480 measured samples;
`comparison.json` retains all 208 comparison rows and 52 favorable changes
above the 5% review threshold. The full default baseline independently passes
two 201-row repeats (6,030 measured samples), generated-catalog/report
validation, and the 38 coverage-index tests. See `acceptance-review.md`,
`gate-summary.json`, and `baseline-capture-summary.json`.

The protocol below describes the original plan. The optional range extension
was not run. `protocol-clarifications.md` and `driver-transition.json` explain
non-interleaved role ordering, supplementary profile retries, and exact
historical helper custody. `verify-final.py` is the final independent checker;
the original `verify.py` and its before-only `verification.json` remain
historical evidence. Failed gate/profile attempts are retained explicitly.

The frozen executables and all raw profile recordings remain in the local
`/tmp/litchi-goal-0501` namespace for replay. That tmpfs retention is ephemeral.
Source archives, build commands, raw reports, and text profile exports are
committed. Strict local verification requires the recorded shared-workspace
source state and retained binary/profile assets; the bundle does not claim a
hermetic clean-checkout binary reproduction for unrelated uncommitted work.
The final verifier can be run with:

```sh
python3 -B docs/performance/results/change-0501/verify-final.py
```

## Original measurement plan

This bundle defines a matched measurement for the private PPTX
`source-backed cross-copy` touched digest. The current planner hashes prepared
image and chart payload bytes into a private plan identity while publication
also performs exact payload and package checks. The candidate may remove only
that redundant digest input after source review proves that plan matching,
source lineage and revision checks, candidate validation, resource admission,
cancellation, and partial-output behavior retain their existing proofs.

The formal matrix has eight serial children: plain and media-rich corpora,
the owned `bytes` provider and a recently written positional `file` provider,
three warmups, thirty retained samples, and two reversed repeats. The `file`
provider is explicitly staged under `/tmp/litchi-goal-0501/corpora`; it is a
warm tmpfs control and carries no cold-storage claim. An opt-in four-child
`range-no-delay` extension uses a 64 KiB caller range cap with zero fixed delay
when the extra run is affordable. It is supplementary and does not replace the
core matrix.

`capture.py` records the exact command, executable and source manifest hashes,
temporary directory and filesystem evidence, report/output hashes, GNU-time
whole-child RSS, and cleanup status. The provider report is independently
checked by `verify-report.py`: every semantic, package, dependency,
relationship, payload, determinism, version, and refusal gate must be true;
source-read histograms and cache/budget phase counters must be monotonic and
arithmetically consistent; and memory, object, and depth reservations must be
zero after the final owner drop. Caller-visible reads are logical adapter
counters, not physical I/O counts.

`compare.py` compares p50, p95, p99, mean, and throughput for API and phase
timers, along with whole-child RSS and all retained source/cache/budget
identity. Every absolute change above five percent is retained in the report
for review. It never turns a flag into an automatic acceptance or discards an
adverse repeat.

`profile.py` runs bounded media-rich owned and warm-file children with `perf
stat` and low-frequency DWARF callchains. Raw `perf.data` is kept only below
`/tmp/litchi-goal-0501`; compact stat/report text and a summary remain in this
bundle. The profile includes setup, corpus gates, warmups, diagnostics, and
serialization, so SHA-256 hits establish hotness evidence for the whole child,
not an operation-local attribution. The baseline profile must be captured
before source edits; an after profile uses the same lane and event list.

The frozen executable is the normal `litchi-perf-baseline` binary built from
`tools/perf-baseline/Cargo.toml` with Rust 1.98.1, locked offline release
settings, disabled incremental compilation and release debug information, and
the dedicated target `/tmp/litchi-goal-0501-target`. Fixture construction,
provider setup, correctness checks, and output publication remain the same in
both roles. This protocol does not claim native-producer breadth, cold or
remote I/O, physical copying, concurrent scaling, operation-local allocation,
or completion of the repository-wide goal.
