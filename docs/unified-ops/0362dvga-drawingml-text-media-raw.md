# 0362dvga — DrawingML text and media raw review repairs

- Date: 2026-10-24.
- Parent: 0362dvg, 0362dv, 0362dvb.
- Status: source repairs complete; coordinator validation is pending a slot.

The independent review found six correctness and admission gaps in the first
validation commit. This repair closes them without changing the default
fidelity profile:

- PPTX `model3d::codec` now has its own required `Limits` policy. Read and
  write map that policy into the shared codec, and package retention passes the
  default policy explicitly. Namespace preparation admits the source before
  any output copy, returns a borrowed source when no declaration is needed,
  checks the completed prepared length before reserving, and uses fallible
  reservation for the bounded copy.
- The shared model3d reader counts the root and every descendant in one
  operation-wide node budget, including direct empty children. Every retained
  child and raster/blip fragment checks its child-count and fragment-byte
  budgets before `to_vec`; raster parsing remains borrowed until that point.
  The standalone opaque path applies the same fragment admission.
- SVG's general `write` preserves unchanged contextual raw fragments by using
  the contextual completion path; a regression test covers exact bytes.
  Canonical size planning now reports the active caller output limit instead
  of failing first against the hard XML ceiling.
- Text-body character data preflights decoded UTF-8 bytes, including entity
  expansion, before `unescape` can allocate. Escaped valid text keeps its
  existing semantic value; the regression test covers an exact four-byte
  decoded scalar and a short budget refusal.

No Cargo command was run while this repair was source-only before the slot.
Rustfmt and `git diff --check` were run before validation. No generated
artifacts or binaries are tracked.

## Coordinator validation receipt

The coordinator granted focused slot 1 after the source freeze. Every command
used the capped cargo shim with `LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`,
`CARGO_INCREMENTAL=0`, offline mode, and `--locked`. The copied lockfile SHA-256
was `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The external target was
`/home/zhuhe/code/litchi-worktrees/.coord/drawingml-text-media-raw/target`
(1.1 GiB after validation); raw logs remain outside the repository.

The focused commands were:

```text
cargo test -p litchi-drawingml --all-features --offline --locked --lib --test svg_blip --no-fail-fast --target-dir /home/zhuhe/code/litchi-worktrees/.coord/drawingml-text-media-raw/target
cargo test -p litchi-pptx --all-features --offline --locked --lib 'model3d::' --no-fail-fast --target-dir /home/zhuhe/code/litchi-worktrees/.coord/drawingml-text-media-raw/target
```

The DrawingML command passed 241 library tests and 24 SVG integration tests.
The PPTX model3d filter passed all 3 targeted tests. A sequential `cargo
check --all-features --offline --locked --lib` passed for
`litchi-spreadsheet-drawing`, `litchi-docx`, `litchi-pptx`, and `litchi-xlsx`.
External receipts are:

| receipt | bytes | SHA-256 |
| --- | ---: | --- |
| `repair-drawingml.log` | 22,091 | `1fa0e16bd981632328d187bb5ace60bedd136f52292f2f8e91fa19aec09cced7` |
| `repair-pptx-model3d.log` | 4,556 | `6217a7009ef6f1e06df1d92b0b0e7a1fa114c07fe6c99ae0dcbcc015f4970952` |
| `repair-consumer-checks.log` | 1,542 | `25586a43785e31a6f266ba8c7b5c5b556474c5bc120e24519664d9a486e11955` |

All validation commands exited zero; no native timing or full-workspace build
was run.
