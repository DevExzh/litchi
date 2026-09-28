# 0809 — restore the baseline PPTX quality gate

Three refusal-test expressions now use `Result::expect_err` directly. This
removes the existing `clippy::err_expect` failures that stopped
[0808](0808-pptx-direct-event-handling.md). Only
`crates/litchi-pptx/src/opened/tests.rs` changes: fixtures, calls, custom
messages, and typed `Error::Invalid`/`Error::Limit` assertions remain intact.
The success types already implement `Debug`; no production type, runtime
behavior, public API, or lint-policy change is required.

The [quality packet](results/change-0809/README.md) records a fresh target,
source identities, exact commands, raw logs, and an independent source review.
Formatting, all-features/all-targets checking, warning-denied Clippy, and
warning-denied rustdoc pass. The full PPTX test run passes **1,238 tests**, with
zero failures and three ignored tests across 85 result groups. The ignored
tests are not claimed as executed. The repository crate-boundary gate also passes, completing all six gates.

The validator reconstructs the exact three replacements from base commit
`679114b237`, checks the complete 9,196-file source census, and replays all
quality receipt and log identities. The ignored root Cargo lockfile and tracked
rustfmt configuration are retained as supplemental inputs captured after the
Cargo gates: their original modification times precede the first gate, all
Cargo dependency commands used `--locked`, and rustfmt bytes match the base
commit. This is documented post-gate provenance, not a retroactive claim that
those inputs were in the original source census.

This is a necessary quality enabler, with no new timing, allocation, RSS,
throughput, or instruction-count result. The archived direct-event candidate
remains unadopted. Resume that experiment with fresh baseline and candidate
builds, retained namespace/error-order obligations, and the declared workflow
benefit and regression gates. No historical timing becomes a paired control.

No new fuzz campaign, native Office, cross-platform, cross-format, or broad CRUD
coverage is claimed. The non-iWork performance program remains incomplete;
iWork is excluded. Unrelated working-tree files remain untouched.
