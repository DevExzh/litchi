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
