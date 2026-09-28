# 0808 results review: early stop at a baseline Clippy failure

The candidate is deferred from this trial. The run stopped at the after
production quality gate before after compilation, paired capture, or profiling
could begin. There is therefore no performance, allocation, profile, or
adoption result to publish.

## Evidence completed before the stop

The retained before leg completed the three instrumented binaries. Probe
quality passed all three gates and recorded 36 passing tests. The before-only
qualification completed 18 reports and 18 samples; its independent
`root-qualification.json` audit accepted the source, output, and semantic
fields exactly and did not use historical timings.

After candidate application, the PPTX quality sequence reached four gates:

| Gate | Result | Evidence |
| --- | --- | --- |
| `cargo fmt --check` | pass | `quality-after/00.log` |
| `cargo check --all-targets` | pass | `quality-after/01.log` |
| `cargo test` | pass: 1,241 passed, 0 failed, 3 ignored | `quality-after/02.log` |
| warning-denied all-targets Clippy | fail, exit 101 | `quality-after/03.log` |

The failed Clippy command reported `clippy::err-expect` at
`crates/litchi-pptx/src/opened/tests.rs:464`, `:538`, and `:557`, with
`expect_err` suggested at each site. The quality driver stopped at this first
failure, so rustdoc and later quality boundaries were not run.

## Why the failure is baseline-owned

The candidate application allowlist contains only
`crates/litchi-pptx/src/notes/codec.rs`. The opened test file is outside that
allowlist and its SHA-256 is
`575d4e0737c52e90a8f062f9fb3d064a2ce58551102254005f8fbfdb37a84253` both in
the recorded base (`d28e3dc702d84f8752a2e329c2d3fc15f5c94f06`) and in the
candidate quality source manifest. The production diff and application
witness likewise contain no `opened/tests.rs` change.

After restoring the source, the same package, feature, target, and
`-D warnings` Clippy command was run as a baseline control. It exited 101 and
reported the same three locations; `baseline-clippy.json` and
`baseline-clippy.log` retain that control. This establishes an existing
baseline hygiene blocker and does not demonstrate a candidate regression.

## Disposition and boundary

The candidate production change was not retained. Root restored the full
`build-before/source.json` manifest, and the archive retains the candidate,
application witness, qualification receipts, probe receipts, and failed
quality logs. The owned target cleanup is separately gated on this restoration
and preserves those records.

No `build-after`, native lane, allocation lane, or profile lane was started.
The 18 qualification samples are a before-only custody check, not a paired
measurement. The candidate remains deferred until the independent baseline
Clippy blocker is repaired and a fresh paired experiment is run.
