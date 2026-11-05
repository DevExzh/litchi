# 0362ul70 — Markdown reader owner-policy boundary

Status: source READY for coordinator review. The focused Markdown profile
passed its guarded validation at source commit
`19767b446796cef6f0df49c48c54ff8bf8528c9e`; this record does not claim final
acceptance or a final all-workspace inventory. The candidate is based on
current integration `8ea056c5dc7283e276a882955de20d71f80377af` in the atomic
branch `uo2/markdown-policy-0362ul70-atomic`. Its verified merge-base is
`ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb` (derived with `git merge-base`,
not guessed).

## Atomic lane binding

The atomic lane is the complete Markdown source, consumer, glossary, and
0362ul70/70a record delta applied to the current integration parent. The
three-way source merge before the final atomic record annotations produced
tree `932af8c9ce5be9d1bdbb51fb07772d041eafbc32`, exactly matching
`git merge-tree --write-tree 8ea056c5dc7283e276a882955de20d71f80377af
a25ab10fa39ee592d33a3520b60d97d251732d6f`; the final atomic tree is checked
against that merge result plus the recorded annotation delta.

The focused validation remains bound to source commit
`19767b446796cef6f0df49c48c54ff8bf8528c9e`. The pre-squash record tip is
`a25ab10fa39ee592d33a3520b60d97d251732d6f`; both refs are preserved locally
as `uo2/markdown-policy-0362ul70-tested-19767` and
`uo2/markdown-policy-0362ul70-pre-squash-a25ab10`. No source or test was
re-run during this atomic preparation.

## Disposition

`litchi-markdown::reader::Snapshot` is an immutable, exact-source owner with
an `Edit`/shared `Commit`/`Patch` lifecycle. It is not a detached raw codec or
record parser. A snapshot retains the complete UTF-8 source in an `Arc<str>`
because source-ranged projections, reversible edits, and exact patch replay
depend on that ownership. The parser validates the configured source, line,
event, block, nesting, and operation ceilings before and during allocation.
The finite defaults on `Snapshot::read` and `Snapshot::read_bytes` therefore
do not imply an unbounded raw API.

The owner disposition keeps those two CommonMark convenience constructors and
adds one canonical retained `OpenOptions` envelope:

* `OpenOptions` selects `dialect` and `read_limits`, with
  `dialect_value()` and `read_limits_value()` accessors.
* `Snapshot::read_with_options` and `read_bytes_with_options` are the
  explicit ingress paths; the selected envelope is retained by the snapshot
  and reused by rendering, edits, and patch replay.
* The former public `Snapshot::read_with(source, dialect, limits)` entry is
  removed. Internal reparsing uses a crate-private `read_with_policy` helper,
  so no public owner entry accepts a per-call limits argument.
* `Snapshot::read` and `read_bytes` continue to use the safe finite defaults,
  as permitted for an owner by 0351 §2.5.3. They are not aliases for an
  unbounded path.

The source `State` now retains the `OpenOptions` value itself rather than
parallel unlabelled policy fields. Every Markdown production/test consumer
that selected a dialect or limits was migrated to the options entry point;
default readers remain unchanged for callers that intentionally select the
default finite policy. The exact-source parser, edit algorithms, patch wire
format, HTML output, and error taxonomy are otherwise unchanged.

## Glossary registration

The new `markdown` profile registers `reader::Snapshot` as an immutable owner,
the `reader` edit family, the shared `reader::Commit` identity, and the
`Snapshot::apply` result contract. It bans the old `read_with` name and public
`*_with_limits` suffixes in this crate. No `required_signatures` raw-limit row
is appropriate: the only public ingress is the owner/options API, while the
parser's limits-bearing function remains crate-private.

## Consumer and validation boundary

The Markdown README, conformance corpus, normative corpus, reader tests, and
all in-crate replay paths use `read_with_options` when a non-default policy is
needed. Focused regressions check text and byte ingress, retained dialect and
limits, committed edits, cross-policy replay, and refusal atomicity.

## Guarded validation receipt

The approved runner was
`/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`
(SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`; helper
`3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f`; Cargo
shim `495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`).
The run used Rust `1.95.0-x86_64-unknown-linux-gnu`, `--offline --locked`,
one Cargo job and test thread, `debug=0`, `incremental=0`, and the 20 GB
decimal target policy with an 18/19/20 GB decimal target start/live/hard
guard (19 GB is the live guard). The hydrated lock hashes were
root `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`,
`tools/perf-baseline` `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`,
and `tools/native-resave`
`51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`.

State and concise receipts are preserved outside the repository at
`(deleted-coord-artifact)`
(target reached 440,957,202 bytes). All seven steps passed: Markdown
all-target tests and Clippy, `litchi-ops` Markdown lib tests/check/Clippy, and
`litchi` Markdown check/Clippy. The seven receipt log SHA-256 values, in that
order, are:

* `915ec469b474276b9d94791a94692b65aaf451293841a356e837476f1164c044`;
* `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` (used
  by the five zero-output steps); and
* `f0c77c7c48b1d222ea3348b4c314a8a6ef1b27b9ad0f4420a105e6625ce24bfc` for
  the `litchi-ops` Markdown lib test.

The earlier failed attempts remain preserved: the unrelated pre-existing
workspace formatting failure is recorded at
`(deleted-coord-artifact)/00-fmt-attempt-markdown-f5-00.log`
(SHA-256
`192f26c38d0ec2aaa45e37b1ddc8c81f6a89a15930f9a19ba8332af5f8a785ac`), the
first compile failure at
`(deleted-coord-artifact)`
(log SHA
`3c3853e0aab07f15fd2e77050b3456ddc9c3bb575911b1ea33fdce9a9ac8dbf0`), and
the subsequent `double_must_use` Clippy failure at
`(deleted-coord-artifact)`
(log SHA
`d8649f6ff233c0431b2a1d65cb4920f2049f8d2f3d3bef4b71eaab26a56216a5`). No
broad all-features operations suite, rustdoc, native benchmark, or final
whole-workspace gate was claimed; coordinator merge and final inventory remain
pending.
