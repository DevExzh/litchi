# 0362utb — ODG ordered replacement state and inherited style dependencies

This successor follows 0362uta without amending it.  The implementation is
`bac9d58fd4`, whose parent is the corrected source commit
`15f677ed5288523f21a396bf4e6f569957fad941` through receipt
`c5e14c8d7f`.

## Ordered replacement state

Joined page-name preflight now follows the payload's ordered
`package.replace` boundaries.  A page-name operation after a replacement is
validated against the replacement package, so a single payload can replace a
package, add a page, and then rename that new page.  Final-name uniqueness and
`draw:page-name` dependency checks remain atomic within each active package
segment.  The original source fingerprint is still required for every durable
operation; only selector and dependency state moves with the active package.

Execution uses the same segmentation.  Operations before the first
replacement stage in the owner edit; operations after a replacement are
accumulated and committed once for that active replacement.  If another
replacement follows, that one semantic segment is committed before the next
replacement is opened.  Security and active-content policies are checked
against the actual current package before each replacement, preserving signed
or active intermediate-state refusals.  There is no per-operation publication
loop; one replacement with semantic follow-up still has one semantic rebuild.

The whole-package replacement commit continues to report byte-level
`changed` state with an empty semantic change list and projection patch.  It
does not fabricate semantic details for an opaque replacement blob.

## Projected style dependencies

Joined style/transition admission now projects page style references in stable
sub-edit identifier order and in each payload's operation order.  A transition
uses the style effective immediately before that operation, including an
earlier style retarget in the same sub-edit.  Comparisons are made only across
distinct sub-edits (or ordinary staged work versus joined work), so an
intentional ordered style-then-transition sequence remains valid.

The dependency test follows drawing-page parent styles with the same bounded
depth, cycle, family filtering, and automatic-style shadowing rules used by
transition writing.  It therefore catches both direct shared styles and a
retarget to a child style whose parent is the transition owner, in either
identifier order.  Replacement boundaries reset the projected page/style
state.  Regressions cover direct and inherited dependencies, ordinary staging
before and after join, and a multi-operation sub-edit whose projected style is
later shared by another sub-edit.

## Validation

The source commit was validated with the memory-capped cargo shim, one job,
`CARGO_INCREMENTAL=0`, debug information disabled, offline locked resolution,
and `TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/odg-joined-v3/tmp`.  The
owned isolated target was
`/home/zhuhe/code/litchi-worktrees/.coord/odg-joined-v2/target`, measured at
414 MiB after the run.  External receipts are intentionally not tracked:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg -j 1` | `.coord/odg-joined-v3/final-check.log` | 214 bytes | `4974a25c9cad8a2363c80e21cb9101f88849defc7901d2c0224ce6dcc40fdb38` |
| `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -j 1 -- -D warnings` | `.coord/odg-joined-v3/final-clippy.log` | 162 bytes | `0cc5a8dc880121e9d8717adce21cf71d18fecd20162df1454c351da73286b868` |
| `cargo test --offline --locked --all-features -p litchi-odg --test package_semantics --test transition_review --test capability_transactions -j 1` | `.coord/odg-joined-v3/final-tests.log` | 4690 bytes | `4af0045aa2531deb0b7a2f27d4b60c10e577f874f637442ed10fdae23cc9d858` |

The focused run passed 8 capability tests, 26 package-semantic tests and 17
transition-review tests (51 total).  `cargo fmt -p litchi-odg -- --check`,
`git diff --check`, the API glossary policy check and the repository artifact
guards passed.  No benchmark, fixture, binary or large repository evidence
artifact was added.
