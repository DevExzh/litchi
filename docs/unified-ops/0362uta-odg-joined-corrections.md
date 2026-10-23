# 0362uta — ODG joined dependency and ordering corrections

This is a successor to the 0362ut source and receipt.  The implementation is
`15f677ed5288523f21a396bf4e6f569957fad941`, a child of the validated 0362ut
line; 0362ut is not amended.

## Corrections

The joined admission and commit paths now preserve the capabilities that were
lost by the first review fix.

- Page-name effects use a page-specific write key.  Independent renames of
  different pages are admitted.  Admission and execution compute the final
  names and refuse duplicate final names atomically, while retaining the
  existing name-dependency checks.  A refused `JoinError` retains the plan and
  leaves the edit unchanged.
- Dependency comparison is between distinct sub-edits.  Operations in one
  prepared sub-edit retain their explicit payload order, so a valid ordered
  page-style/transition sequence is not rejected merely because the two
  operations touch the same page.
- A style retarget is compared with the transition's source style, not only
  with its page.  Thus a page-0 retarget to a style currently used by page 1
  conflicts with page-1 transition work in either identifier order.  The same
  dependency is checked when ordinary staged work surrounds a joined plan;
  genuinely independent styles remain composable.
- A single durable payload may intentionally order `package.replace` and
  semantic operations, as the standalone durable API permits.  Independent
  sub-edits that mix package replacement and semantic effects remain
  conflicting.  The joined path validates the replacement policies, reopens
  it once, stages the later semantic operations into one edit, and performs
  one semantic commit; it does not publish through a per-operation loop.

The source-fingerprint check, typed refusals, composition bounds, deterministic
sub-edit order, replay/inverse behavior and private payload boundary remain in
force.  The added regressions cover distinct page renames, duplicate final
names, an ordered style/transition payload with replay/inverse/durable proof,
cross-page shared-style dependencies in both orders, ordinary-before/after
joined dependency checks, and the ordered mixed replacement capability.

## Validation

The source commit was validated with the memory-capped cargo shim, one job,
`CARGO_INCREMENTAL=0`, debug information disabled, offline locked resolution,
and `TMPDIR=/tmp/claude-1001/odg-joined-v2`.  The isolated target was
`/home/zhuhe/code/litchi-worktrees/.coord/odg-joined-v2/target`, measured at
414 MiB after the run.  Receipts remain outside the repository:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg` | `.coord/odg-joined-v2/check.log` | 3891 bytes | `1888ab6e7c314961b38cba479d89802625ff1c92be39030a5d1b81a5b3de07f7` |
| `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -- -D warnings` | `.coord/odg-joined-v2/clippy.log` | 162 bytes | `b8097b628cac4c2c60c31d3503f4cf35f5ca1af7360cadae9abababefa159f78` |
| `cargo test --offline --locked --all-features -p litchi-odg --test package_semantics --test transition_review --test capability_transactions` | `.coord/odg-joined-v2/tests.log` | 4539 bytes | `f7b36d9e258afcfc5edb477d2e4575407673410b448a0ca2f97e70c921b87c45` |

The focused tests passed 26 package-semantic tests, 14 transition tests and 8
capability tests (48 total).  Rustfmt, `git diff --check`, the API glossary
policy check and the repository artifact guards passed.  No benchmark,
fixture, binary or large repository evidence artifact was added.
