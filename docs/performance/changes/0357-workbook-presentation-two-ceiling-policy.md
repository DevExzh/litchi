# Change 0357: workbook and presentation two-ceiling path policy

Status: implemented

`performance_claim: none`

## Two-ceiling filesystem path policy

Workbook and Presentation filesystem paths now distinguish caller-limited
OOXML candidates from finite neutral fallback owners. ZIP/OOXML and uncertain
or polyglot candidates use the caller's `ReadLimits`, capped by a neutral
2 GiB absolute ceiling. Ordinary canonical ODP/ODS inputs, content-derived
suffix-renamed ODP inputs, OLE inputs, and generic non-ZIP inputs use the
neutral 2 GiB policy. Unknown or missing-content-types ZIPs and ODF inputs
with a present or uncertain OOXML catalog remain caller-limited, preserving
the caller's resource policy. Explicit byte and typed source APIs remain
strict and are unchanged.

## Single-source presentation arbitration

PPTX, ODP, native PPT, and the bounded `Bytes` fallback now arbitrate from one
`FileSource` and one `SourceVersion` on filesystem paths. Owner probes share
the pinned source, and freshness is checked without pathname reopening or an
unbounded read. Content-derived canonical and renamed ODP inputs stay with
the ODP owner; terminal `OtherOoxml` and `DisabledOtherOoxml` outcomes do not
fall through to another owner. The workbook path applies the same two-ceiling
distinction while retaining its format-owned source behavior.

The ODF catalog detector has a neutral-budget helper that applies the checked
2 GiB input, compressed-byte, entry-count, and total-byte ceilings together.
This keeps ordinary ODF fallback finite while leaving caller limits in force
for uncertain/polyglot arbitration.

## Regression coverage

The regressions cover canonical and renamed ODP ownership, ordinary ODS
fallback, OOXML/polyglot caller limits, exact neutral-ceiling failures,
wrong-family terminal outcomes, one-source freshness, generic non-ZIP and
native-PPT fallback, and the ODF neutral catalog budget. The focused
`litchi-odf-common` `detect::tests` run passed `15/15` with `260` filtered;
the `litchi` `catalog_detection_arbitration` integration test passed `6/6`;
the `litchi` `pptx,odp,ppt` library run passed `82/82`; and the `litchi`
`ods,xlsx` library run passed `84/84`.

## Validation and resource boundary

All validation ran serially with `CARGO_BUILD_JOBS=1`,
`CARGO_INCREMENTAL=0`, `CARGO_PROFILE_DEV_DEBUG=0`,
`CARGO_PROFILE_TEST_DEBUG=0`, `ulimit -v 8388608`, one disk target at
`/home/zhuhe/CodeProjects/.cargo-targets/change-0357`, and
`--test-threads=1` where applicable:

```sh
cargo test -p litchi-odf-common --lib detect::tests -- --test-threads=1
cargo test -p litchi --no-default-features --features pptx,odp,ppt --test catalog_detection_arbitration -- --test-threads=1
cargo test -p litchi --no-default-features --features pptx,odp,ppt --lib -- --test-threads=1
cargo test -p litchi --no-default-features --features ods,xlsx --lib -- --test-threads=1
cargo check -p litchi --no-default-features --features pptx --lib --tests --quiet
cargo check -p litchi --no-default-features --features odp,ppt --lib --tests --quiet
cargo check -p litchi --no-default-features --features odp --lib --tests --quiet
cargo check -p litchi --no-default-features --features ppt --lib --tests --quiet
cargo check -p litchi --no-default-features --features xls,xlsx --lib --tests --quiet
```

The initial constrained compile exposed two `Arc<FileSource>` to
`Arc<dyn ReadAt>` coercion errors; both were corrected before the final
checks. The target's final/peak observed footprint was 1.3 GiB. Host
availability was approximately 14 GiB with 133 GiB disk free and swap
exhausted. No parallel build was used and no OOM occurred. These are
correctness/resource observations only; there is no speed, RSS, allocation,
constant-memory, or OOM-prevention claim.

## Remaining scope

The public eager `DetectedFormat` API remains. A neutral 2 GiB fallback can
still fully materialize its input. Flat ODF MIME decoding can allocate before
the strict MIME bound. Aggregate Presentation `Vec`/`join` paths remain
infallible. Portable same-size replacement identity is not fully provable;
native PPT probe-time mutation coverage is incomplete; and the `Current User`
plus `Workbook` OLE classifier inconsistency remains. Prepared ODP package
reparse remains where applicable. `litchi-opc::parts_by_name` case lookup and
selected-Part materialization remain open.
