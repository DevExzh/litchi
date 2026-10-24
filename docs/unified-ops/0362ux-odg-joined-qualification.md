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

`tools/odg_joined_qualification_builder.py` is the paired build/preparation
entry point. It copies
`tools/perf-baseline/odg_joined_qualification_common.rs` to
`crates/litchi-odg/examples/odg_joined_qualification_common.rs`, and copies
`tools/perf-baseline/odg_joined_qualification_allocator.rs` to its paired
allocator module, then copies
exactly one of the two adapters to
`crates/litchi-odg/examples/litchi-perf-odg-joined-qualification.rs`. The
manifest records each overlay source digest and destination, the clean
worktree's immutable commit SHA and tree SHA, the complete `git ls-tree`
source manifest, an actual byte closure of every tracked blob (including stat
identities for cheap between-probe checks), `Cargo.lock`, `.cargo/config.toml`,
and `rust-toolchain.toml` digests, the enabled feature list, and both normal and
allocator binary identities. The baseline is admitted only at pristine
`ad52437f5e5e2c71ea476eedd5ddcf61aee2021d`; the candidate ref is bound to the
final candidate field. The builder rejects sparse/skip-worktree omissions,
symlinked paths, and `3rdparty` inputs. The runner requires
the exact paired binary basename and rejects stale refs, dirty trees, changed
lock/toolchain/probe/fixture files, or changed binaries.

The overlay is a build-staging input, not a source mutation accepted by the
runner: the external build receipt retains its source/destination digests,
streams each Cargo child into exclusive bounded logs with live RSS, output,
load, process, and target guards, then removes the temporary example files
and verifies both source worktrees are clean before the execution manifest is
resolved. Preparation and build receipts, targets, logs, and runtime output
must all be fresh; receipt validation rehashes every log and artifact. The
builder invokes the virtual workspace with the bound
--package litchi-odg --release --example
litchi-perf-odg-joined-qualification command and admits only the corresponding
target/release/examples/ artifact. The manifest carries external identities
for the normal and allocator binaries and the completed build receipt; a
self-edited receipt cannot establish those identities.

The adapter and fixture paths are fixed to the reviewed repository files and
their approved bytes, rather than being caller-selected look-alikes. Both
git ls-files -v and git ls-files --debug are checked, so assume-unchanged,
skip-worktree, extended-index, sparse, missing, or symlinked inputs fail
closed. Cargo.toml is admitted alongside Cargo.lock, rust-toolchain.toml, and
.cargo/config.toml, and its actual bytes are recorded before and after each
build event. Every build command records the complete effective environment,
resolved Cargo/rustc paths and hashes, ancestor/home Cargo and rustup config
identities, and the endpoint cwd identity; nonempty unreviewed compiler
selection (`RUSTC`/`CARGO_BUILD_RUSTC`) and wrapper variables are rejected.
Runtime phases reuse the cached closure and
perform cheap stat probes between launches, with full byte revalidation at
phase boundaries so stat-hidden edits fail closed. The final target byte
sample is checked against the 20 GB hard cap both at execution and receipt
validation.

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
calibration, and fixed `a1,b1,b2,a2` ABBA ordering. It requires exactly two
ABBA blocks, 20 warmups, and one sample per fresh child, a 500 ms calibration
target, and a 150 ms minimum accepted sample. Each child report and log is
bound to endpoint, resolved ref, binary digest, route, scope, phase, live
guard, target, and rehashed stream identities; malformed, extra, stale, or
incomplete rows fail closed.

Calibration uses the explicit reviewed adaptive-doubling-v1 policy: the
reported attempts must start at one iteration, double up to the declared
maximum, and bind the final count and elapsed time. Before every build,
allocator launch, calibration, ABBA leg, and temporary-directory use, the
runner applies the 19.2 GB preflight and 20 GB hard bounds. Directory scans
are no-follow scans which reject missing entries, symlinks, and disappearing
files. The child guard installs a parent-death signal and a Linux child
subreaper, records owned PID/start-time/process-group identities, streams
bounded output, and rejects detached descendants after the child exits.

`--allocation-only` runs the same routes/scopes and ABBA order against a
separate allocator-instrumented build. The overlay allocator publishes
operation-scoped global-system-allocator counters whose region starts before
durable plan reconstruction and closes after the selected apply/publication
boundary. Allocation reports carry no latency claim and validate live-byte
conservation, peak bounds, tool identity, and the fixed one-sample envelope.

The runner writes only to a new external output directory. Use, for example:

```text
python3 tools/odg_joined_qualification.py \
  --manifest /external/0362ux/manifest.json \
  --out-dir /external/0362ux/run-v1 \
  --correctness-only
```

Build preparation and admission are explicit and receipt-bound:

```text
python3 tools/odg_joined_qualification.py --manifest /external/0362ux/manifest.json \
  --out-dir /external/0362ux/unused --prepare
python3 tools/odg_joined_qualification.py --manifest /external/0362ux/manifest.json \
  --out-dir /external/0362ux/unused --build
python3 tools/odg_joined_qualification.py --manifest /external/0362ux/manifest.json \
  --out-dir /external/0362ux/unused --validate-build
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
python3 -m py_compile tools/odg_joined_qualification.py \
  tools/odg_joined_qualification_builder.py \
  tools/odg_joined_qualification_bindings.py tools/test_odg_joined_qualification.py
python3 -m unittest -v tools.test_odg_joined_qualification
rustfmt --check tools/perf-baseline/odg_joined_qualification_common.rs \
  tools/perf-baseline/odg_joined_qualification_allocator.rs \
  tools/perf-baseline/odg_joined_qualification_probe.rs \
  tools/perf-baseline/odg_joined_qualification_ad524_probe.rs
```
