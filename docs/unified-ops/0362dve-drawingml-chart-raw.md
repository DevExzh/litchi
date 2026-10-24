# 0362dve — DrawingML chart raw codec limits

- Date: 2026-10-24.
- Parent source checkpoint: `2f7db7ee7fa01893e9099217fc4243687d196f63`.
- Status: source migration, consumer migration, and capped validation complete;
  the packet is ready for coordinator review.

## Contract

The chart raw codecs now share one caller-selected
`litchi_drawingml::chart::Limits` policy. Its default preserves the established
ceilings and its fields cover source/output bytes, classic-chart XML nodes,
depth, attributes and attribute values, authored series/categories/points and
text, ChartEx bytes/nodes/depth, and chart-style bytes/nodes/depth/attributes,
strings, modifiers, colors, variations and transforms. Every policy is
validated before parsing or writing; invalid values return the typed
`Error::InvalidLimit`, while observed exhaustion returns `Error::Limit` with
the actual `LimitKind`.

The required raw entry points are:

```text
chart::reader::read(reader, &Limits)
chart::authoring::read(xml, &Limits)
authoring::Parsed::is_authored(xml, &Limits)
chart::chartex::read(xml, &Limits)
chart::style::{parse, parse_color}(xml, &Limits)
chart::writer::{write, write_with_rels}(..., &Limits, ...)
```

The classic reader bounds its input and MCE output before materializing the
typed chart, counts XML nodes/depth/attributes while streaming, admits series,
category points, numeric points and semantic text before retaining them, and
validates the completed chart. ChartEx and chart-style tree builders apply the
caller's node/depth/attribute/string/semantic bounds before proportional
vectors or owned values are built. ChartEx authoring escapes directly into a
bounded formatter, and the classic writer uses a counting sink, so oversized
output is refused before the caller's sink receives bytes beyond its policy;
neither path builds an unbounded intermediate output buffer.

DOCX and XLSX pass their retained chart and companion-part policies through
the shared limits type. PPTX, spreadsheet-drawing, OPS, DrawingML tests and
examples use the canonical default where no narrower host policy is retained.
An exact tracked-consumer search covered `crates/`, `tools/`,
`docs/report/spec-gap-validation-evidence/`, tests, examples, fuzz and
performance manifests. The report harnesses contain no live chart raw reader,
authoring, Chartex, style or writer call site (the chart-formatcode2 harness
is a separate extension codec); historical receipts and source pins were left
unchanged.

## Validation

The workspace lock fingerprint used by every Cargo command was:

```text
Cargo.lock SHA-256: 52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6
Cargo.lock git blob: ce39fdc781093aa0819bcd17d41af891d4f33bcc
```

The source checks were `rustfmt --edition 2024` on all changed Rust files and
`git diff --check`. With the repository capped shim, one build job, offline
locked dependencies, one test thread, zero incremental compilation, an
external `TMPDIR`, and external target
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-chart-raw-target`
(5,977,097,991 bytes, below 20 GB), these commands passed:

```text
cargo test --locked --offline -p litchi-drawingml --all-features --lib --tests --no-fail-fast --quiet
  238 unit tests plus every integration target passed; chart_limits: 8 passed
cargo test --locked --offline -p litchi-drawingml --all-features --test chart_limits --no-fail-fast --quiet
  8 passed
cargo check --locked --offline -p litchi-pptx -p litchi-spreadsheet-drawing -p litchi-xlsx -p litchi-ops --all-features
  passed (the DOCX chart dependency was built in this consumer check)
cargo test --locked --offline -p litchi-docx --all-features --lib chart::tests --no-fail-fast --quiet
  9 passed
```

The consumer check and DOCX chart test ran after the public signature migration;
the final DML-only rerun additionally covered the late-admission hardening for
Chartex attributes/text, chart categories, style modifiers, generated
scatter/bubble arrays and bounded escaping. No native timing was run in this
lane, no Cargo process remains active, and no generated binaries or large
performance evidence are committed.
