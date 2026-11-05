# 0362ul158 — PPT hyperlink source-limit repair

Date: 2026-10-25. Parent: 0362ul154b. Status: the production repair is
committed on the leaf branch; coordinator merge and the composed final gate
remain pending.

## Scope and source binding

The leaf is
`uo2/ppt-hyperlink-crud-repair` in
`/home/zhuhe/code/litchi-worktrees/uo2-ppt-hyperlink-crud-repair`. It was cut
from the frozen C source
`f8cd398a770f7363f628be40edb3ce4219566c16`. The production repair is
`72baeeb13d22358e578e262d11b81bda3b5ccbae`,
`fix(ppt): bound source-backed OfficeArt shape parsing`. Its binary source
diff from C has SHA-256
`44d259a79e11059851b502cc5275a167e82e84b6cf1e679c951b99cb72a1ecb1`.

The external setup packet is
`(deleted-coord-artifact)`.
It binds the leaf source and execution environment without adding build
artifacts to this repository:

| input | bytes | SHA-256 |
|---|---:|---|
| `launch.json` | 1,129 | `bfa4c1d48a31d335e6b47624dcb6825ca5aa3e273f487592cdb8cd16e8b66662` |
| `stdout.log` | 120 | `82f70e0b36fa8569076d0c9e1022ee70bd7117d264d5bf09addc67b3f3029d04` |
| `stderr.log` | 0 | `e3b0c44298fc1c149af4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `setup-events.jsonl` | 778 | `075afff8ab1ef758d516a843162888f14a8e6f5065d315512235fb961914effb` |
| rules file | 1,848 | `ee1a985cd31c004d4b3422b7e834b419d1e646cfd8ff1abda1bdac203a870b89` |
| `env.sh` | 474 | `c1b6600669694c3b703e75ea9cf8aa2d293fc26e18a345e93b8014f95d5d5ccc` |
| capped Cargo shim | 520 | `1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8` |

The setup launch bound the process group to `MemoryMax=4G`,
`MemorySwapMax=0`, and `KillMode=control-group`. The author's focused Cargo
commands used the leaf environment, offline/locked resolution, one build job
and one test thread. They were direct capped-Cargo invocations rather than
`run_safe_gate.run_step` executions, so their terminal observations are not
accepted gate evidence. The external target was
`(deleted-coord-artifact)`;
its observed size after the runs was 3,249,787,719 bytes (3.25 GB), below the
20 GB target bound. A post-run process check found no `cargo`, `rustc`,
`crud_matrix` or `litchi-ppt` process.

## Failure on C

The C1 full gate stopped at row 25, the PPT hyperlink add operation. The
retained failure log is

`(deleted-coord-artifact)/0005.log`,
376,777 bytes, SHA-256
`235b2e984bb00d71b6aaffee237e20f9764fc822f023fdd0850298deae6727fd`.
The corresponding terminal verification receipt is 6,769 bytes, SHA-256
`5d8b95e578bc1f687082657f3a6357c59808617c1077f37843e0b1eb7b3b6d15`.
The failed gate command was:

```text
cargo test -p litchi-ops --all-features \
  --test crud_matrix --test doc_unified_ops --no-fail-fast \
  --message-format json-render-diagnostics
```

It reported `corrupt_file` / “the file is malformed” at PPT hyperlink add.
The six later hyperlink cells were therefore not reached; this was not accepted
as a skipped capability. The C source and its locks were unchanged by the
failed run.

## Diagnosis

The failing call path is:

```text
set_shape_hyperlink
  -> slide_order::elements::shape_click
  -> Slide::shape_interactions_with_limits
  -> odraw::limits_for_source
  -> litchi-odraw shape parsing
  -> ExObjList hyperlink publication
```

The PPT outer `SourceLimits` default allows a record-tree depth of 256.
OfficeArt's format-neutral shape parser has its existing public safety policy
`litchi_odraw::Limits::default().max_depth == 64`. Passing the PPT profile
unchanged to that parser made the parser reject an ordinary source-backed
drawing before the hyperlink publication was reached. This explains both the
malformed-file error and the six unreached cells.

The focused diagnostic session was 95140. Its error chain was:

```text
PPT native edit operation failed
OfficeArt error: OfficeArt Depth limit 256 exceeds the safe maximum of 64
OfficeArt error: OfficeArt Depth limit 256 exceeds the safe maximum of 64
OfficeArt Depth limit 256 exceeds the safe maximum of 64
```

The initial all-feature CRUD reproduction was session 45557 and reproduced the
C failure on clean C. The fixed all-feature CRUD run was session 96046 and
was observed to complete with `unreached claimed cells (0)`. Exact commands
and source bindings for these author observations are preserved in the
reduced external ledger described below; they are not a substitute for the
coordinator's approved runner receipt.

## Repair

The helper now intersects the caller's source profile with OfficeArt's own
public safety policy:

```rust
let defaults = litchi_odraw::Limits::default();
max_depth: source
    .max_depth()
    .min(usize::from(defaults.max_depth))
    .min(u16::MAX as usize) as u16,
```

It does not duplicate the value 64. A caller-selected smaller depth remains
effective. The source `max_record_bytes` and `max_events` are retained
(the latter maps to OfficeArt `max_records`). The source-derived profile is
now passed by every affected source-backed shape projection, including
animations, programmable tags, shape flags, placeholders, text-edit source
resolution, and the hyperlink path. Record-only callers continue to use the
explicit neutral default.

The production commit changes these seven tracked files:

- `crates/litchi-ppt/src/odraw/codec.rs`
- `crates/litchi-ppt/src/odraw/tests.rs`
- `crates/litchi-ppt/src/slide/types/codec.rs`
- `crates/litchi-ppt/src/slide/types/semantic.rs`
- `crates/litchi-ppt/src/slide/types/tests.rs`
- `crates/litchi-ppt/src/text_edit/source_range.rs`
- `crates/litchi-ppt/tests/element_hyperlinks.rs`

The 0294 regression
`elements_exobjlist_hyperlink_add_reopens_with_chart_and_content_intact`
adds a shape hyperlink, reopens the result, preserves all pre-existing
hyperlinks, checks the chart inventory and identity, checks unchanged slide
text, and applies the inverse patch to prove exact source-byte restoration.
It is a fidelity guard, not a capability waiver.

The test correction in the follow-up record commit strengthens the helper
regression. The test now parses a real `DgContainer` containing a `Dg`
OfficeArt record, so it reaches the shape parser. It also constructs a
restrictive source profile with `max_record_bytes=1024`,
`max_events=16`, and `max_depth=1`, then verifies all three values survive
the projection and that the drawing is admitted. A nested `SpgrContainer`
then exceeds that same caller depth and must be rejected, proving that the
source-selected ceiling remains behaviorally restrictive. The prior
empty-input assertion could only prove that an empty loop returned an empty
vector and could not exercise OfficeArt depth admission.

## Validation

The external reduced ledger is intentionally not a raw Cargo log or an
accepted runner receipt:
`(deleted-coord-artifact)/repair-158-receipt.json`,
4,183 bytes, SHA-256
`6a19ae530f3efc59c3b17900b6d5aacb24f51615196bf3356caa9017da0f7955`.
It records the exact argv, source state, session identifiers, and the author's
observations; every run is marked `accepted: false`, and raw stdout was not
retained. The leaf did not produce the required `run_safe_gate.run_step`
log/source/process receipt. The coordinator must rerun the focused checks on
the committed leaf source through that approved runner before accepting the
repair.

The author's focused observations (pending approved rerun) are:

| check | result |
|---|---|
| clean-C all-feature `crud_checklist_per_format` reproduction | expected PPT hyperlink failure; six later cells unreached |
| diagnostic hyperlink error-chain test | identified OfficeArt depth 256 versus safe 64 |
| repaired all-feature `crud_checklist_per_format` | pass; zero unreached claimed cells |
| `cargo test -p litchi-ppt --test element_hyperlinks -- --nocapture` | 3 passed, including the 0294 regression |
| corrected real-OfficeArt limit unit | 1 passed |
| outline-resolution corpus-slide regression | passed |
| `cargo fmt --all -- --check` and `git diff --check` | passed in the earlier leaf validation; coordinator rerun remains authoritative |

The corrected unit command observed before the final nested-depth assertion
was:

```text
source (deleted-coord-artifact) &&
cargo test -p litchi-ppt --lib \
  odraw::tests::source_limits_intersect_officeart_safety_ceiling \
  -- --exact --nocapture
```

The all-feature CRUD command was:

```text
source (deleted-coord-artifact) &&
cargo test -p litchi-ops --all-features --test crud_matrix \
  crud_checklist_per_format -- --exact --nocapture
```

The full CRUD command was not rerun after the helper-test-only correction:
production Rust behavior was unchanged from the already passing 72bae source
commit, and the correction only replaced an empty parse input with a real
OfficeArt record and added restrictive-profile assertions. The nested-depth
assertion is likewise test-only and still requires the approved focused
rerun. The coordinator should not treat the direct-Cargo 96046 observation as
a runner receipt; run the focused test and any composed final gate on the
merged source.

## Fidelity, performance and boundaries

The fix restores an already intended source-backed operation; it does not
remove any operation, weaken the parser's OfficeArt safety ceiling, or change
the caller's restrictive limits. The 0294 regression covers hyperlink
retention, chart fidelity, text fidelity and exact inverse bytes. No ABBA
measurement was performed in this focused repair lane. The helper performs
only scalar limit intersection and default-policy construction; no measured
performance claim is made, and final composed performance qualification
remains pending.

The tracked 0294 fixture is
`test-data/unified-ops/0294/elements.ppt`, 11,776 bytes, SHA-256
`245ead4cce6d1cafc1cb619e16f12b6d8e2856daab90c0478e94cec987e4a262`.
The leaf also verified the PPT fixtures
`test-data/ole/ppt/SampleShow.ppt` (SHA-256
`64275a6685b8f267827178037dcc59de927475e56385725916f4c6039c7ce7f6`) and
`test-data/ole/ppt/empty.ppt` (SHA-256
`cc2bd402a951c443d6296b855c49ad79710da2c6ffb16db3ebde55351621ae6a`).
No `3rdparty/` path, benchmark binary, large evidence file, or capability
waiver was added.

The external test binaries and their identities remain outside Git; these
sizes and hashes are author observations from the direct runs, not accepted
runner evidence:

- `crud_matrix-61d367eaa518655e`: 258,543,784 bytes, SHA-256
  `e64b076503f50c05e48e0350e1f0207ba48e5eb78d5a02949d85e6c1efb791d5`
- `element_hyperlinks-0fb1f7c3ad60a4b9`: 10,511,872 bytes, SHA-256
  `27ff70a994cc2041fecef449a883140a897173c60cf1c7e68704ad9ec7f5db8`
- `litchi_ppt-467504ad8775dece`: SHA-256
  `bf9a07f827a84e20044d3eb2a36e2777b71a8e3e4ef6ca949ed1521c96a1222a`

All large evidence and binaries remain external. Nothing is pushed by this leaf.
