# 0362utf — ODG lazy edit XML parts

This successor does not amend the prior ODG receipts.  The implementation is
`05c15847c1`, whose parent is `08afcb24bc`.

## O(1) edit construction

`Package::edit_with_policies` now stores private `LazyPart` states for
`content.xml` and optional `styles.xml`.  Each state retains only an O(1)
`Package` clone and borrows the source part until its first real mutation;
`set_owned` installs the produced replacement string at that mutation.  The
helper has no derived debug representation, so source and staged XML remain
outside public/debug payloads.

Read-only selectors and policy checks continue through `Deref` without part
copies.  No-op comparison uses borrowed strings and does not clone
`styles.xml`.  `insert_page` uses `Cow` for its temporary styled content, so a
page without a generated transition style does not clone `content.xml`.
Pure joined package replacement reaches the replacement owner without
materializing either XML part.

The source unit proof checks source-pointer retention for both parts, dropping
an untouched edit, materialization only after a real page mutation, exact
no-op commit behavior, and an applied pure replacement with both lazy parts
still unmaterialized before publication.

## Validation

Validation used the memory-capped cargo shim, one job,
`CARGO_INCREMENTAL=0`, debug information disabled, offline locked resolution,
and `TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/odg-joined-v4/tmp`.
The isolated target was
`/home/zhuhe/code/litchi-worktrees/.coord/odg-joined-v2/target`, measured at
590 MiB after the run.  External receipts are intentionally not tracked:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg -j 1` | `.coord/odg-joined-v4/y-final-check.log` | 162 bytes | `e9b1edd644f4f94cdf7f1a03a51418306eaccde04d9c26c741e5760f2aecd26e` |
| `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -j 1 -- -D warnings` | `.coord/odg-joined-v4/y-final-clippy.log` | 162 bytes | `d62823c17e4c55a83207940903cf2908c652ff3628990b8fafdac5af9b90fac5` |
| `cargo test --offline --locked --all-features -p litchi-odg -j 1` | `.coord/odg-joined-v4/x-full-suite.log` | 18746 bytes | `40ef24f598952523fa82a4f58b24c8e4ac9095f1e878bcec58aa535defe8f839` |

The full suite passed 23 library tests, 155 integration tests, and 1
doctest (179 total).  The focused lazy unit receipt was also successful
(`x-lazy-unit.log`, 506 bytes, SHA-256
`efe7eb3b221eaddf94fa64786a8dac00195e112d9e54dd4857c18e71e90f3081`).
`cargo fmt -p litchi-odg -- --check`, `git diff --check`, the checked API
glossary policy parse (55 profiles), and the repository artifact guards
passed.  No benchmark, fixture, binary, or large repository evidence artifact
was added.
