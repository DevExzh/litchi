# 0362hda — ODF flat output admission follow-up

- Date: 2026-10-22.
- Parent: 0362hd (`90d88df9d8`, based on `8d6abdab69`).
- Status: validated source follow-up; coordinator integration and the final
  workspace inventory remain required.

The first flat raw-limits slice had three correctness gaps. `BoundedBytes`
reserved the difference between its target and capacity even though
`try_reserve_exact` measures additional bytes from the current length. A
large write after a spare-capacity prefix could therefore make the final
`Vec` grow infallibly beyond the sink bound. The reservation now uses the
current length, with a regression that writes 196 bytes after a four-byte
prefix in a 64-byte-capacity, 256-byte-bounded sink and proves the capacity
stays bounded.

The XML scanner now checks the attempted child depth for `Event::Empty` as it
does for `Event::Start`, including the exact next-depth refusal. `Parts`
retains the maximum depth admitted by its split and rechecks the proof in
`to_bytes` and `merge`, so a package split under a loose profile cannot be
published through an unchanged route under a tighter depth profile. The
output paths also recheck the retained flat source/content size admission.

`Parts` now retains a SHA-256 content proof of its source rather than another
full source buffer. `merge` verifies the digest after the cheap length check
and returns the typed `Error::PatchConflict` for a foreign source, including
a same-length source whose bytes differ. This closes the old range-reuse
behavior while retaining the bounded source memory model. The digest is
computed once at split and once per merge; no source copy is retained.

The synthetic ZIP projection has an explicit `Limits::projection_options`
carrier: it uses the hard synthetic source scratch ceiling while retaining the
caller's semantic limits. ODT, ODP, ODS and ODG projection consumers now pass
that carrier when reopening generated ZIP bytes. This preserves exact flat
input ceilings without incorrectly applying them to a larger internal ZIP.

Focused and affected validation completed with the capped environment from
`.coord/w8-odf-flat-limits/env.sh`, `CARGO_BUILD_JOBS=1`, and offline mode:

* `cargo test -p litchi-odf-common --all-features --offline flat::tests` —
  11 passed.
* `cargo test -p litchi-odt -p litchi-ods -p litchi-odp -p litchi-odg
  --all-features --offline --lib` — 807 + 694 + 221 + 22 passed.
* `cargo test -p litchi-ops --all-features --offline --lib` — 965 passed.
  An earlier run at the same command stopped at 964 passed and one failure:
  `tests::document::odt::flat_odt_source_limits_remain_typed`; that exposed
  the projection reopen using the original flat source ceiling and was fixed
  by the consumer migration above.
* `cargo clippy -p litchi-odf-common -p litchi-odt -p litchi-ods -p
  litchi-odp -p litchi-odg -p litchi-ops --all-targets --all-features
  --offline -- -D warnings` — passed.
* `cargo fmt --all -- --check` and `git diff --check` — passed.
* The exact glossary rustdoc command that previously exposed the iWork-only
  detector warning was rerun and passed:
  `RUSTC_BOOTSTRAP=1 cargo rustdoc --package litchi --lib --target-dir
  /home/zhuhe/code/litchi-worktrees/.coord/w8-odf-flat-limits/iwork-rustdoc
  --no-default-features --features iwork --offline -- -Zunstable-options
  --output-format json --document-hidden-items`.

The matched comparison is prepared against source `8d6abdab69`: split the
same tracked flat fixtures, publish a package, merge an exact no-op and a
single-child replacement, and compare exact bytes, refusal kinds, and
fallible-reservation paths. The source digest and retained-depth checks are
covered by the focused tests; any native timing or allocator qualification is
intentionally deferred until the coordinator grants a quiet measurement
window. Target size was 4.6 GiB at handoff; no benchmark binaries or large
evidence files were added.
