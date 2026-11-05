# 0362dvg — DrawingML text and media raw limits

- Date: 2026-10-24.
- Parent: 0351, 0362dv, 0362dvb.
- Status: source migration and focused validation complete; coordinator integration remains.

This lane gives the remaining shared DrawingML raw codecs one required policy
at every resource-scaled boundary. Text-body reads now use the single
`reader::read(xml, scope, limits)` entry; text-body writes use
`writer::{write,write_contents}(..., writer::Limits)` and count the complete
output before reserving or changing the caller's string. Their existing
`Builder` and fixed-domain scalar APIs remain unchanged.

`model3d::codec::{read,opaque,write,write_to}` now take
`model3d::codec::Limits`. The read policy bounds complete input, opaque
fragments, child counts, raster children, depth, and nested nodes. Writers
preflight the serialized size; `write_to` then emits directly to the caller's
sink rather than materializing a second complete buffer. Invalid policies use
`InvalidLimit`, and over-budget failures retain the typed `LimitKind` plus
observed and maximum values.

SVG blips now expose one top-level API with `svg_blip::Limits`; the former
nested codec is crate-private. Standalone and contextual reads, vector writes,
and both sink writes all require the policy. Contextual output performs the
existing size preflight before allocation or the first sink write. The
arbitrary-XML `blip::find_first_embed` scanner likewise requires its typed
input/relationship policy and checks input admission before constructing the
reader.

PPTX, DOCX, XLSX, spreadsheet-drawing, DrawingML tests, and the live XLSX SVG
allocation harness were migrated in this source change. Historical baseline
and matched evidence snapshots under `docs/report/spec-gap-validation-evidence`
were left byte-for-byte unchanged. The glossary now bans the removed
`text::body::reader::read_with_options` spelling in favor of the canonical
required-limits entry.

No Cargo command was run in this worker before the coordinator grants the
validation slot. `rustfmt`, `git diff --check`, and glossary JSON parsing were
run locally. No generated binaries, third-party paths, or large evidence
artifacts are part of this commit.

## Coordinator validation receipt

The coordinator granted one focused validation slot after the source change.
All commands below used the capped cargo shim with `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, offline mode, and `--locked`.
The copied lockfile was pinned to SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The external target directory was
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-text-media-raw/target`
(691 MiB after the checks), and no target or raw log is tracked here.

The first attempt compiled dependencies but stopped on the source-only
`unreachable_pub` visibility lint and an unused formatter import. Those were
fixed in the follow-up source commit before the passing rerun. The canonical
DrawingML command was:

```text
cargo test -p litchi-drawingml --all-features --offline --locked --lib --test svg_blip --no-fail-fast --target-dir /home/zhuhe/code/litchi-worktrees/.coord/drawingml-text-media-raw/target
```

It passed 238 library tests and 22 `svg_blip` integration tests. The consumer
command was run sequentially for `litchi-spreadsheet-drawing`, `litchi-docx`,
`litchi-pptx`, and `litchi-xlsx`, each with `--all-features --offline
--locked --lib`; all four checks passed. External receipts are:

| receipt | bytes | SHA-256 |
| --- | ---: | --- |
| `drawingml-focused-rerun.log` | 21,737 | `1daeec36d1c1ccacb4a6af580a666eca1e12e2269faed8c75b2a61cf69a55b00` |
| `consumer-checks.log` | 12,817 | `cec7c9af1ecb8b519a08e257fe5c9e1b89318113f915ea7bcd6ce2bcbd7f8db7` |

The failed first-attempt receipt remains external at
`drawingml-focused.log` (6,941 bytes,
`6cb6f638a20b2e2d6f19ae2fbdcc6e55820a6bc372fb510ad54fbe44698c965a`) for
traceability. The validation commands exited zero after the repairs.
