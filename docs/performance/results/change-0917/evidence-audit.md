# Change 0917 DOC evidence and reproduction audit

This is a read-only audit of `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`.
It excludes iWork. I did not build, run a native capture, modify production
code, or commit anything. The only file owned by this audit is this note.

## Fixture and source binding

The two real DOC inputs are unchanged at this checkout:

| case | path | bytes | SHA-256 |
| --- | --- | ---: | --- |
| `docnohf` | `test-data/ole/doc/NoHeadFoot.doc` | 26,112 | `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48` |
| `docfloat` | `test-data/ole/doc/FloatingPictures.doc` | 335,360 | `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d` |

The 0730 adoption commit `4599c1225da767986e28633e3b332de6223e9881` is an
ancestor of this HEAD, but the packet was built from an older candidate. The
current source hashes are therefore part of a fresh qualification identity:

```text
crates/litchi-doc/src/body_text.rs                 0e8cb6351be3327322882ae4035230cd241aed0b6208cb8589e7d76692d3d786
crates/litchi-doc/src/tracked_revision/package.rs  a04da643b093ec9724c9ac57ed8fbfdf2c2fedfc16eade105c91c32df7c9a446
crates/litchi-ole-common/src/object/editor.rs      ee6079567bf85a0ff68a335fb43999d780fab335a09a9cec8b845c16b7b5b319
```

`docs/performance/results/change-0730/candidate-source.json` records different
post-change hashes (`c1447266…`, `ee9672af…`, and `4f01d8…`). Its binaries are
also historical: `docs/performance/results/change-0730/cleanup.json` records
both `doc_retention_probe` and `doc_retention_probe_alloc` as removed (their
old SHA-256 values were `37b91c7a…` and `f536e071…`). No retained 0730
executable was found under `/home/zhuhe`. Old 0728/0729/0730 timing values
must not be used as current timing evidence.

The 0906 legacy audit is likewise historical and explicitly says it ran no
fresh timing; its recommendation for a bounded DOC handoff is superseded by
the adopted 0730 handoff. It remains useful as a coverage warning only:
registered selectors and synthetic cases do not prove real-producer edit
coverage (`docs/performance/0906-office-coverage-qualification.md:3-24`,
`docs/performance/results/change-0906/legacy-coverage.md:1-16`).

## Current runnable route

The closest reusable harness is the candidate-only 0730 retention probe at
`docs/performance/results/change-0730/retention-probe/`. Its public route is
the right ownership path:

```text
Snapshot::open_bounded(source.to_vec(), Limits::default(), transaction_limits)
  -> snapshot.edit()
  -> Edit::replace_paragraph(Position::new(0), replacement)
  -> Edit::commit()
  -> commit.snapshot().bytes().to_vec()
```

The route and allocator boundary are in
`docs/performance/results/change-0730/retention-probe/src/lib.rs:199-289` and
`alloc_metrics.rs:87-139`. At this read-only audit snapshot, the working copy
at `docs/performance/results/change-0917/probe/` was the same adaptation point,
but its parser and result storage still accepted only `one`/`two` and used a
two-element stage array. Any later extension must preserve the route and
allocator boundaries described here.

The future many-edit source should accept one count per process from
`1,2,8,32`, use the two fixed 0730 texts alternately, and target paragraph zero
in the same public `Edit`. Alternation makes every requested operation a real
change after the first; the probe must still assert that the first replacement
differs from the source and that `patch.changes().len() == count`. Counts 32
fit the current default transaction bounds (256 operations and 4 MiB aggregate
replacement units) with these short texts.

Use the same transaction bounds for all controls and vary only retention:

```rust
TransactionLimits::default().with_max_retained_render_bytes(0)             // zero
TransactionLimits::default()                                                // default, 8 MiB
TransactionLimits::default() /* plus release after every successful stage */ // release
```

`TransactionLimits::new` is not the zero control for this comparison because
it also changes the operation and replacement limits. The current APIs that
make this behavior explicit are `TransactionLimits` and its default at
`crates/litchi-doc/src/body_text.rs:50-135`,
`Edit::retained_render_bytes`/`release_retained_render` at `:1513-1524`, and
the lower owner’s capacity filter at
`crates/litchi-doc/src/tracked_revision/package.rs:1128-1159,3020-3055`.

After the external probe source has been extended, the reproducible build/run
shape is:

```sh
CARGO_TARGET_DIR=/home/zhuhe/litchi-performance-0917/target \
  cargo build --manifest-path docs/performance/results/change-0917/probe/Cargo.toml \
  --release --offline --locked --bins

taskset -c 12 \
  /home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe \
  --case docnohf --input test-data/ole/doc/NoHeadFoot.doc \
  --edits 32 --retention default \
  > /home/zhuhe/litchi-performance-0917/docnohf-32-default.json

taskset -c 12 \
  /home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe_alloc \
  --case docfloat --input test-data/ole/doc/FloatingPictures.doc \
  --edits 8 --retention release \
  > /home/zhuhe/litchi-performance-0917/docfloat-8-release-alloc.json
```

The full matrix is two fixtures × four counts × three controls. The frozen
0917 protocol schedules six rotated ordinary rounds (144 child processes, three
warmups and 30 samples per child) and two allocator repetitions (48 child
processes, one measured owner each). Binaries, targets, raw JSON, stderr and
logs belong under `/home/zhuhe/litchi-performance-0917`, outside Git.

`tools/perf-baseline` has a separate current source route, but it is not a
real-file replacement harness. `doc_owner_public_phases` dispatches to
`run_doc_owner_public_phases` (`tools/perf-baseline/src/lib.rs:25263-25264,
:27689`) and always uses generated `doc_fresh_write_to` writer shapes. The
documented command is useful only as a synthetic diagnostic control:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --writer-shape tiny \
  --case doc_owner_public_phases --json /tmp/doc-owner-public-phases.json
```

Its `--ole2-file` facade selectors only open/read named files; they do not add
the requested real-DOC edit/commit route. Do not substitute this synthetic
case for the two fixture matrix.

## Qualification gates and preservation oracle

The measured changed lifecycle should start immediately before the input clone
and public open and stop after commit and the final output `Vec` copy. Source
file I/O, fixed text preparation, expected/reference output, CFB inventory,
semantic readback, patch checks, JSON serialization and process setup stay
outside the timer. Optional `release_retained_render()` calls belong inside the
changed lifecycle because they are part of the selected route.

Before capture, run a functional qualification for every fixture/count/control
and require all of the following:

* Bind source bytes, size and SHA-256; verify every stage is a distinct
  paragraph-0 change and the final text is the last requested text.
* Reopen source and output through the public DOC reader. Compare paragraph
  zero and all surviving paragraphs, paragraph count, main-story projections,
  stories, fields, revisions, table cells and embedded-object collections. The
  0730 semantic witness in `docs/performance/results/change-0730/probe/src/lib.rs`
  is the starting oracle, with its one-edit expectation changed to the final
  many-edit text.
* Check changed commit state (`changed()`, non-noop patch, count of semantic
  changes), then apply the forward patch to the original and an independently
  reopened identical source. Apply `patch.inverse()` to the changed snapshot
  and require exact source bytes. Applying the forward patch to the changed
  snapshot and to a different valid fixture must return the current conflict
  refusal. These APIs and exact-source rules are at
  `crates/litchi-doc/src/body_text.rs:2327-2445`.
* Run a separate exact no-op route for each retention policy. Require unchanged
  source bytes, `changed() == false`, `patch.is_noop()`, zero changes, and no
  retained render. If diagnostics are enabled, require the current no-op event
  sequence (`Finish`, `ExactNoOp`, `Patch`) outside the changed timing samples.
  `Edit::commit`’s no-op branch is at `body_text.rs:2086-2107`.
* Exercise typed refusal and failure atomicity outside the timed matrix:
  out-of-range paragraph → `Refusal::ParagraphNotFound`; malformed source →
  open refusal; over-limit replacement/operation → its typed limit refusal;
  source/editor state and retained token unchanged after each failed mutation.
* Reuse the 0730 CFB oracle (`probe/src/lib.rs:523-818,1458-1537,2273-2415`)
  against the zero-retention many-edit output, then compare default and release
  output bytes to that control. Require valid CFB source/control/output, equal
  stream path sets, no changed stream outside `WordDocument`, `0Table`, and
  `1Table`, byte-identical untouched streams, preserved root/storage CLSIDs,
  unchanged directory shape and permitted allocation-field-only raw-directory
  differences. Do not accept output equality alone as the preservation proof.
* Check source policy on every route: the opened snapshot and committed
  snapshot retain the selected `TransactionLimits`; release changes token
  ownership only; and patch application uses the componentwise intersection of
  source/before/after policies. The implementation points are
  `body_text.rs:853-857,1337-1341,2412-2435`.

The lower common editor already returns the validated candidate rendering to the
DOC owner (`crates/litchi-ole-common/src/object/editor.rs:457-486,509-555`).
The final public reopen and source retention therefore remain mandatory for
every stage and commit; retention is an ownership optimization, not a skipped
validation path.

## Allocation and observer interpretation

The allocator binary reports process-local, region-relative
`allocated_bytes`, `deallocated_bytes`, allocation calls, `peak_live_bytes`,
and boundary `retained_bytes`. It does not report RSS, total process memory,
phase-specific allocation, or physical I/O. The region should start before
`source.to_vec()`/public open and end after the output `Vec` is returned. Read
the fixture before region entry. Use a fixed stack array such as
`[StageObservation; 32]`; returning a heap `Vec<StageObservation>` would make
its bookkeeping appear as retained bytes and invalidate the protocol’s
“output allocation only” interpretation.

`peak_live_bytes` is the maximum live allocation delta above the region-entry
baseline. Default retention can lower repeated-render allocation while raising
transient peak live bytes: a clone-first mutation drops the old token only as
the new candidate is published, so old and new allocations can overlap. A
successful retained token is a `Vec` capacity capped at 8 MiB, not serialized
output length. At region end the token is moved into commit and the probe keeps
only the copied output `Vec`; hence `retained_bytes` should describe the output
allocation boundary, not prove that the handoff stayed resident. Release is
expected to resemble zero retention in allocation fields, but this is a
measurement to verify for all four counts, not an old 0730 assumption.

If semantic observer timing is needed for qualification, use a separate
`performance-diagnostics` build with
`Snapshot::open_bounded_profiled` and `Edit::commit_profiled` from
`body_text.rs:700-807,2110-2215`. Open events are strict-owner validation,
public-reader validation and source retention; changed commit events are
finish, strict-owner validation, public-reader validation, source retention and
patch; an exact no-op has finish, exact-no-op and patch. The observer is
synchronous, so callback/recorder cost is inside the outer interval. It does
not cover `replace_paragraph` staging, and nested CFB events from the separate
`*_with_cfb_observer` API must not be added to semantic phase totals. Keep this
observer pass separate from ordinary lifecycle timing; the feature and
callback/lifetime path are different. No 0729 or 0730 phase percentage should
be inferred for the current HEAD.

The largest evidence gap is therefore a current-source, real-fixture,
many-edit matrix that combines the adopted public handoff with complete
preservation/patch-policy qualification and allocator peak-live observations.
The archived 0730 probe proves only one/two edits, the current synthetic phase
harness cannot consume these real DOC edits, and the old executables are gone.
Until the adapted probe is built from the current source and the fresh matrix
passes these gates, there is no current timing or memory conclusion to report.
