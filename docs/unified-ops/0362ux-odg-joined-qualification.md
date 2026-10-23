# 0362ux — paired ODG joined qualification tooling

This record adds the source-specific qualification harness for the ODG joined
composition migration. It is tooling and protocol only; it contains no native
build, timing, or qualification result.

## Endpoint boundary

The baseline probe is built only in a pristine checkout at
`ad52437f5e5e2c71ea476eedd5ddcf61aee2021d` and calls
`Package::apply_joined(JoinedEdits)`. The candidate probe is built only in the
post-migration checkout and calls `Package::edit`, `Edit::join`, and
`Edit::commit`. They are separate source adapters so the baseline cannot
silently acquire the candidate commit triad.

The external build overlay copies
`tools/perf-baseline/odg_joined_qualification_common.rs` to
`crates/litchi-odg/examples/odg_joined_qualification_common.rs`, and copies
exactly one of the two adapters to
`crates/litchi-odg/examples/litchi-perf-odg-joined-qualification.rs`. The
manifest records each overlay source digest and destination, the clean
worktree's immutable commit SHA and tree SHA, `Cargo.lock` and
`rust-toolchain.toml` digests, the enabled feature list, and the external
binary identity. The runner requires
the exact paired binary basename and rejects stale refs, dirty trees, changed
lock/toolchain/probe/fixture files, or changed binaries.

The overlay is a build-staging input, not a source mutation accepted by the
runner: the external build receipt retains its source/destination digests,
then removes the temporary example files and verifies both source worktrees
are clean before the execution manifest is resolved.

## Routes and clocks

Every endpoint proves all four routes in untimed mode:

* `package_replace_one` — one real package replacement;
* `semantic_two_disjoint` — shape text and name on disjoint shapes;
* `semantic_four_disjoint` — four disjoint semantic changes;
* `empty_join` — exact no-op.

Each route has `join_apply` and `join_apply_publish` phase rows. Source
opening, fixture assembly, edit preparation, output digest validation, and
correctness/refusal proofs are outside the clock. The timed closure starts
before bounded durable-plan reconstruction, including every allocation needed
to rebuild the `JoinedEdits`; its per-sub-edit wire input and count are
bounded by the fixed composition limits. Publication and readback are inside
the second scope.

The candidate untimed proof additionally checks the returned diagnostics,
direct patch replay, durable forward/inverse replay, and forward replay on a
content-matching package. Both endpoints prove source immutability, exact
no-op/source bytes, deterministic output, semantic values, stale-lineage
refusal, and conflicting-effect refusal. The runner cross-checks source,
route output, publication, and semantic proofs between endpoints without
crediting the baseline with candidate-only capabilities.

## Execution protocol

`tools/odg_joined_qualification.py` reuses the generic native qualification
helpers for manifest JSON, executable identity, Git checks, external target
guards, `/proc` process checks, load checks, CPU-pinned `taskset` launches,
calibration, and fixed `a1,b1,b2,a2` ABBA ordering. It requires a 500 ms
calibration target and 150 ms minimum accepted sample. Each child report and
log is bound to endpoint, resolved ref, binary digest, route, scope, and
phase; malformed, extra, stale, or incomplete rows fail closed. Allocator
capture remains an external metric collected by the shared native
qualification workflow and is not mixed into elapsed-operation claims.

The runner writes only to a new external output directory. Use, for example:

```text
python3 tools/odg_joined_qualification.py \
  --manifest /external/0362ux/manifest.json \
  --out-dir /external/0362ux/run-v1 \
  --correctness-only
```

The full calibration/ABBA command is the same invocation without
`--correctness-only`; it must be run only after both immutable binaries pass
the correctness phase on a quiet host. No result is implied by this source
commit.

## Source and fixture

`test-data/odg/joined-content.xml` is a small deterministic fixture generated
for this protocol with `PackageWriter`; it was not copied from a third-party
checkout. `test-data/odg/README.md` records its provenance and license. The
adapter and runner contract tests cover duplicate manifest keys, altered
manifest fields, stale source refs, lock/toolchain/probe/fixture identity,
binary/log mismatch, false route proofs, incomplete ABBA legs, fixed limits,
and the calibration floor.

Static validation for this source-only commit is:

```text
python3 -m py_compile tools/odg_joined_qualification.py tools/test_odg_joined_qualification.py
python3 -m unittest -v tools.test_odg_joined_qualification
rustfmt --check tools/perf-baseline/odg_joined_qualification_common.rs \
  tools/perf-baseline/odg_joined_qualification_probe.rs \
  tools/perf-baseline/odg_joined_qualification_ad524_probe.rs
```
