# 0362ul51 — SpreadsheetDrawing raw boundaries

This leaf carries the SpreadsheetDrawing raw-boundary migration from the
integration base `267ea715266ffc3bd754bfbfa4896b139adad9fc`.

## Public boundary

The shared `litchi_spreadsheet_drawing` crate now exposes typed module
policies. The variable raw entry points require the policy at the call site:

```text
objects::{read, chart_frames, count, next_id, append, check_picture,
          picture_fragment, shape_fragment, change, detach}(..., &objects::Limits)
shape::read(xml, &shape::Limits)
vml::shapes(xml, &vml::Limits)
```

`objects::Limits`, `shape::Limits`, and `vml::Limits` retain the established
hard ceilings and can be derived from the owning OPC or XLSB source policy.
Append counts every top-level fragment element and admits the post-append
element and drawing-byte totals before allocating the result. Change rewrites
borrow the selected source span and measure each replacement under the
remaining output budget before growing a replacement string; the fragment
writers do not build an oversized temporary string before refusal. Image
relationship sizes are checked during object resolution. Shape reads forward
the selected geometry policy to `CustomGeometryBuilder`, and the shared text
reader retains the selected shape policy without falling back to a global
default. VML scans admit both source bytes and retained shape count before
retention.

`Emitter::for_objects` validates the complete shape/group/connection graph,
including nested object counts and selected custom-geometry dimensions, before
reserving any drawing IDs. The shared geometry writer exposes a bounded
measurement pass, so the selected custom-geometry output budget is admitted
before recursion, reservation, or surrounding XML emission. Top-level object
count remains independent from the aggregate nested-object budget, allowing a
small number of groups to contain many valid children without consuming the
top-level quota.

The XLSX and XLSB drawing, comments, workbook-load, and authoring consumers
all pass their owner-derived policy. Shared limit and invalid-policy failures
remain typed through both format error conversions, including the shape and
VML dimensions.

## Source checks

Focused boundary tests cover selected shape bytes, custom-geometry guide
limits, exact/one-under geometry output admission for both preallocated and
incremental emitters, failure-atomic incremental writes, multi-element append
admission, bounded object rewrites including no-op and shrink cases,
pre-reservation graph validation, invalid object policy, and VML input/shape
limits. The glossary registers minimal and full defining-crate profiles with
the required raw signatures for objects, shapes, VML, and the geometry-policy
writer constructors.

## Bounded validation receipt

This receipt records the leaf's initial diagnostic runs. They are not f5
acceptance receipts: they used the approved memory-capped Cargo shim directly,
without the required ownership and live-target guards. The initial report
mistakenly treated an exited PID as evidence that the runner was unavailable;
root verified the runner file remained available and required the guarded
rerun recorded below.

The final passing steps ran with source tip `5beaa11ceb` frozen (the earlier
diagnostic steps are retained as repairs were made); the integration base was
`267ea715266ffc3bd754bfbfa4896b139adad9fc`. The pinned lock was copied only
to the ignored worktree `Cargo.lock`, with SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`; it was
removed after the run. Rust `1.95.0`, `--offline --locked`,
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, debug info off, and incremental
compilation off were used. The fresh external target was
`(deleted-coord-artifact)`,
which measured 3,986,141,931 bytes at capture (<20 GiB). Logs are retained
outside the repository at
`(deleted-coord-artifact)`.

Successful steps (all exit 0) were:

* `cargo fmt --all -- --check` (`01-fmt.log`);
* DrawingML library, geometry, and transform tests (`02`–`04`);
* SpreadsheetDrawing all-target tests, 42 unit tests plus 2 fixture tests
  (`34-spreadsheet-drawing-all-targets-final.log`);
* XLSX and XLSB all-target checks (`35-xlsx-check-final.log`,
  `36-xlsb-check-final.log`) and all-target tests (`13`'s affected tests
  rerun individually in `14`–`25`, and `26-xlsb-test-all-targets.log`);
* DrawingML library clippy and SpreadsheetDrawing/XLSX/XLSB all-target
  clippy with `-D warnings` (`31` and `32`);
* `litchi-ops` XLSX/XLSB all-target check (`33-ops-xlsx-xlsb-check-all-targets.log`).

The first full XLSX all-target test log (`13-xlsx-test-all-targets.log`) is
retained as a diagnostic failure: 1,879 tests passed, while nine targets
reported `tempfile` `NotFound` errors after the shared TMPDIR disappeared.
The cause of that disappearance remains unknown because the diagnostic run
had no filesystem observer. The affected tests all passed with the fresh sibling TMP
path in logs `14`–`25`. Earlier failed logs (`05`, `08`, `10`, and `27`–`30`)
record compile/lint repairs made in this leaf; the foreign DrawingML diagram
test lints seen in `30` were not changed. No native timings or f5 acceptance
claims are made from this receipt.

Key external log hashes (the log directory is the authoritative full set):

| step | bytes | SHA-256 |
| --- | ---: | --- |
| 01 fmt | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| 02 DrawingML lib | 26,734 | `76a6578a4f75d31e79bcab2df2b6ed88431dcf5f6204e27197b71603d1e7aa54` |
| 03 geometry | 553 | `0653592b0529a4f77161a66f2d4f5e96f9cd0bdc25f670a6c01718e85f5adc60` |
| 04 transform | 501 | `077ac7359d9161716f8c035a447640af600eef0019327713eca2b4a5c7d1ef6e` |
| 13 XLSX diagnostic | 274,399 | `f2b151a7de0e534b219930ed7935e4a6fe329143c3135f32e77f31b81023e88e` |
| 14–25 XLSX reruns | 439/443/428/409/423/429/399/400/699/404/435/396 | hashes below |
| 26 XLSB all-targets | 124,498 | `8b37da8e1a3e8fd0a341f91951b26d3c0c97d879b2718087dcd00c9ddff5cbca` |
| 31 DrawingML clippy | 1,432 | `d4c21c9c79212e0eef2fb5bcf7315d8cc8707b54d64fe9b4ad8a48c8841ff770` |
| 32 Spreadsheet/XLSX/XLSB clippy | 60 | `2a7f2bc267ab84a783d3e33a26c82086964c4f5a781be885b3a698fdde636801` |
| 33 ops check | 2,312 | `b7ce91351a99d6435ef5c658ee9cbb14b9a109b92ac950e6609558c838e3c3be` |
| 34 SpreadsheetDrawing final | 4,118 | `c33fc1a94b9b78eb3be5cbcf0a062102b06631a9ebf7b6f74c2b034277406827` |
| 35 XLSX check final | 321 | `91176bc32cfccc630ed3d019511787b77b7405f4383e476c8764329c29cc646d` |
| 36 XLSB check final | 321 | `3988f1a081779d0cf189759e72d4d25c3d4d63a3d01051b3857dd85c687d2fe6` |

The individual XLSX rerun hashes are, in order, `14`
`3e67a2b0e491df579016a7e50135a35ebdec7a03013f02c31b4a1572087fc532`,
`15` `2736ac29be8a65122d6c8b3ca09972d608b2e55d2a5ecf8b212cf85515cdae8b`,
`16` `0229a236c5b25d30d325b2456744c8e45b0eddfe45898ee2d544458d276dbf01`,
`17` `fc7f503fbfe1db7791d32848ec310ec70a9c9fc17a7ec7720c61f9762f576b86`,
`18` `3884b666067e58d1bb48685766534411d7a680669d35c9d66b56832255e6b33a`,
`19` `660fb65115ae720719250f0f3467e05068941095271c9348954dcf296652a36e`,
`20` `5a5de3ba2bde89e3ce9bb2dc0922af6eb72f6328b9714e88b226717b47c2d433`,
`21` `a247451fdb389d548a88e305b294aac1c62f35a2643de313506546ccd80d3106`,
`22` `a953c231ecc7fd6239c9ae2b1bcd0362f43295518302fa7e76d9e3fd4815c9ff`,
`23` `dabd5e251ba20b12d4c2a381f9831d3e40d2836c8e4d30ffab89fe3c0106edc`,
`24` `8290515818de0d42424d6010d6a95c25b139cc07206f10721f9310af5b46bb72`,
and `25` `c1f703f66aeeb6e8ae06b64cc8c03657a5d60ba0af09313ea0e25b4ba00b4c35`.

## Coordinator review and guarded validation

Atomic `9f43932dd0a5a9f5aef4ce439915f1563f2e6940` passes independent
review of the final delta from `e71f7b95e2`. It differs from diagnostic
`5beaa11cebb5a6e832f45b71fd86ee46c9db294f` only in this record. Owner-derived
profiles, exact output admission, geometry preflight and typed format errors
remain intact. No direct standalone performance-harness consumer of these
SpreadsheetDrawing raw functions was found.

All eleven actual f5 steps pass on clean `9f43932dd0`: formatting;
DrawingML library, geometry and transform tests; SpreadsheetDrawing all-target
tests; XLSX and XLSB all-target checks; XLSX/XLSB ops check; DrawingML library
Clippy; SpreadsheetDrawing/XLSX/XLSB all-target Clippy; and ops Clippy.
The external state is
`(deleted-coord-artifact)/state.json` (70,630 bytes, SHA-256
`45de0b60dc09d9a205b0d296192e9e0cbc5f01ddd79d0b3b05ac858235ad472c`).
Its driver log SHA-256 is
`c085ed2f2c03ee8e0059d1f91fa032fb28628c2eb9647f7b44358a2c6ab91af7`.
The runner SHA-256 is
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.

Root verifies every command-log and Cargo-trace identity, unchanged source
and three locks, empty terminal owned-process lists, and final target size
826,790,824 bytes. The earlier broad Clippy attempt's seven diagram lints and
the intentionally terminated broad test attempt remain retained; neither is
counted as a passing row. This scoped result does not replace the integrated
gate or actual native/performance qualification.
