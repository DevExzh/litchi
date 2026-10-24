# 0362uge — XLS lazy edit entry

This change removes the package-payload staging cost from the XLS edit entry
points covered by 0362uge. `Package::edit` now retains the admitted owner and
policy handles and creates the cell snapshot on the first source lookup,
staged verb, join, or commit. The comments, sheet-visibility, chart,
pivot-view, OLE-object, and User Names package families use the same deferred
owner shape. Source-backed cells, comments, and worksheet visibility retain
their dynamic source owner and materialize the staging package only when a verb
or commit needs it.

The deferred state is private. A first operation still performs the complete
format validation under the selected limits, and subsequent operations reuse
that parsed owner. Signed/encrypted policy checks, source fences, exact no-op
publication, patch/replay behavior, and bounded output paths remain at their
existing boundaries. User Names projection accessors are fallible and
materialize through the same path rather than hiding a parse error behind a
panic. Chart lookup now propagates a deferred-open error instead of converting
it to `None`.

The cell transaction also rejects a second scalar value at a coordinate that
already has an authored Formula staged in the same transaction; this restores
the existing failure-atomic overlap contract.

## Validation

Source tip before this change: `7559f44aa975254bd96f6e28eee17366b0ebeb6a`.

The isolated Cargo target was
`/home/zhuhe/code/litchi-worktrees/.coord/xls-lazy-entry/target` (1.4 GiB at
the last check); no binaries or captures are committed. Commands used with
the assigned one-job offline/locked environment were:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-lazy-entry/env.sh
CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 cargo test -p litchi-xls --lib --no-default-features --locked --offline
CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 cargo test -p litchi-xls --lib --no-default-features --locked --offline user_names::tests
CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 cargo test -p litchi-xls --test xls_user_names --no-default-features --locked --offline
CARGO_INCREMENTAL=0 CARGO_BUILD_JOBS=1 cargo test -p litchi-xls --no-default-features --locked --offline --test xls_charts --test xls_comment_transactions --test xls_ole_objects --test xls_form_controls --test xls_pivot_view_reader --test package_owner --test xls_user_names
```

The User Names unit and integration runs passed 10/10 and 16/16. The first
full library run reached 1,318/1,319: it exposed the pre-existing staged
Formula/scalar overlap assertion, which is fixed in this change and passed in
the isolated rerun. The focused integration run compiled and reached
`package_owner` 21/23 before the deferred-entry assertions were updated to
expect first-operation materialization and commit-time limit refusal. A fresh
rerun of that final focused command is required when the coordinator returns a
Cargo slot; this record makes no native performance or qualification claim.

No benchmark or native ABBA run was performed.

## Review follow-up

The deferred source-owner review added the remaining freshness boundaries in
this successor: comments and worksheet visibility now fence the original
source before planning, after plan validation, and after deferred publication;
their committed owners retain the deferred overlay instead of reopening it at
commit time. The chart and pivot editors restore their private lazy state when
capture or parsing fails, and chart `reorder_at` opens the deferred editor
before inspecting its workbook. The primary package edit/replay guard now
retains the original dynamic source across `Package::from_source` handoff, so
an original mutation is reported as typed `SourceChanged` rather than target
I/O. Focused source-backed tests cover comment/visibility commit, write,
replay and inverse freshness plus primary handoff edit/replay.

The broad all-target XLS suite was first started under the capped target but
intentionally interrupted when these review blockers were reported; that
partial external log remains diagnostic only. The final isolated suite,
Clippy, and affected facade consumer checks are recorded below. No native
timing or qualification claim is made here.

The follow-up also keeps the recipient's original source fence when an
equal-content cross-copy patch is rebased, and primary package guards
re-check that original source if a composed-target version query fails. The
public source-backed commit/replay regressions assert the documented
`litchi_xls::Error::Io` boundary; output and read projections retain their
typed source-change assertions. Chart and pivot retry tests verify a failed
lazy capture does not consume the editor's private source state.

## Final isolated validation

The final source tip is `b8ed2d46bbd06e2434b5b829b594698ed80fdb89`.
The admission regression is an internal test over a dynamic composed target;
it verifies that an in-flight target error is promoted to the retained
original owner's typed `SourceChanged` refusal. It passed 1/1. The focused
integration run passed `package_owner` 23/23 and `source_backed` 84/84.
The all-target/all-feature XLS run passed 1,870 tests in 100 test suites, with
one intentionally ignored test. Strict all-target/all-feature Clippy passed;
the `litchi-ops` all-target/all-feature check passed; and the XLS OLE metadata
consumer test passed 4/4.

Commands, using the assigned one-job capped offline/locked environment:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-lazy-entry/env.sh
cargo test -p litchi-xls --lib package::edit::tests::capture_source_guard_promotes_target_error_to_original_source_changed --locked --offline
cargo test -p litchi-xls --all-features --test source_backed --test package_owner --locked --offline
cargo test -p litchi-xls --all-targets --all-features --locked --offline
cargo clippy -p litchi-xls --all-targets --all-features --locked --offline -- -D warnings
cargo check -p litchi-ops --all-targets --all-features --locked --offline
cargo test -p litchi-ops --all-features --test ole_metadata --locked --offline
```

External receipt logs (not committed) and SHA-256:

```text
admission-unit-b8ed.log         418 B     1373bd43bc1b26dece7323288772e14859e13a30972ce56d13221bf71b848a04
focused-source-owner-b8ed.log   8,385 B  da96b55d28b4bf293d73baf8e2b917319bb1ddd0440479f74dd1322b3ea1b676
full-xls-suite-b8ed.log         169,179 B 5a972ee4e8b2c5f43146ffaff70d1a201a27d13c21ca47df900a44d818a12359
clippy-xls-all-targets-b8ed.log 269 B     0f2a6da43584c091d0e1e0847463e7adf1bd62c37964700e8a1110f7dc1e6455
check-ops-all-targets-b8ed.log  269 B     2fa2e5e506cda33bd5977cf7717de30ab0a9d2fcf03710b2cd6ff672e8427eb2
ops-xls-ole-metadata-b8ed.log   784 B     acb4f39cba2b76d8ff32beb77f404143c19eee81b425493e4bb817721329d8ab
```

The target directory was `/home/zhuhe/code/litchi-worktrees/.coord/xls-lazy-entry/target` and measured 4,553,717,719 bytes after validation. No benchmark, native ABBA, or qualification claim is made.
