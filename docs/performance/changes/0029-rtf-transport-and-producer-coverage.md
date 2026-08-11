# RTF transport and producer coverage

Date: 2026-08-11

Production base: `f7e102ace3c19bcf71f31514ef134f3ec2145ca6`

Disposition: coverage and reproducibility tranche accepted. No RTF production
code, public API, dependency boundary or iWork/IWA crate changes in this
tranche. This record makes no latency or memory improvement claim.

## Coverage change

The seven existing semantic RTF case names now accept an explicit
`--rtf-variant` selector. `plain` remains the default, so historical commands,
the 36-case/198-record default matrix and the 110 selectable case-name count do
not change. Multi-variant result identity is explicit in
`configuration.rtf_variants`, `corpus.rtf_variant`, the variant-bearing corpus
name, shape and archive digest.

| Variant | Input contract | Timed capabilities |
|---|---|---|
| `plain` | Existing deterministic direct ASCII corpus | Open, list, one paragraph, full text, exact stream, exact no-op, one paragraph edit/save |
| `byte1252` | Deterministic `\ansicpg1252` bytes with literal `0xe9`, opened through `Document::from_bytes` | Open, list, one paragraph, full text, exact stream and exact no-op |
| `lzfu` | Deterministic public `transport::compress` output with exact decompression proof | Open, list, one paragraph, full text, exact compressed stream and exact no-op |
| `watermark` | Content-addressed `test-data/rtf/watermark.rtf` producer fixture | Tiny selector only: open, list, one paragraph, full text, exact stream and exact no-op, with three header shapes and `gtextUNICODE=ASAP` required |

Changed byte-1252 body publication is excluded because candidate validation
currently refuses the raw-byte layout after a staged change. Changed LZFu
publication is excluded by the explicit transport-aware-rewrite refusal.
Header watermark shape editing is not exposed by the ordinary body editor.
These capability filters are tested directly; unsupported pairs cannot be run
by bypassing the CLI loop.

All four variants selected at tiny produce 25 rows: 7 plain, 6 byte-1252, 6
LZFu and 6 watermark. Tiny plus large produces 44 rows: 14 plain, 12
byte-1252, 12 LZFu and the one six-row tiny watermark slice. Push CI exercises
the 25-row matrix; scheduled/manual release CI exercises 44 rows and checks
unique `(case, variant, shape)` identities.

## Deterministic evidence

The checked release smoke contains one sample for every permitted tiny row:
[`rtf-variant-coverage-smoke.json`](../results/rtf-variant-coverage-smoke.json),
SHA-256
`10ec4813e464b64b3325b1bf434fa8870dc04eaafb9bc3c8436037751f9cf2f7`.

| Tiny corpus | Bytes | SHA-256 |
|---|---:|---|
| Plain | 1,347 | `ee4a5c5b5d1c97d5fb4f1e862c2787a859136b237addd0d14a7d52ddc9e62328` |
| Raw CP-1252 | 1,407 | `47a20904dfb8107bb1cd9ad099decfed13c76cbde993fdd93eda3d919a9bb1aa` |
| LZFu | 294 | `bf755db7d4afc26a66ffab476884431e6e585f3259df5b6469e2d4fadfc51baf` |
| Producer watermark | 69,471 | `48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb` |

The public-facade correctness target separately pins the MS-OXRTFCP reference
frame, literal CP-1252 exact/no-op behavior and atomic changed-publication
refusal, the LibreOffice watermark themed header shape, and the checked
`relsize` chain:

- source: 516 bytes,
  `1a079582281767c1bf7afa5ef2e63553400cdbc4704aa25d9dbcc34e2c22569d`;
- Litchi-changed: 634 bytes,
  `d0bf70e50972bbf15dc9b0da96b9702d64a92a676cb81529bf28729a2cd91d71`;
- LibreOffice-resaved: 4,051 bytes,
  `224707aea42c7b38712bc66a76424d3666b52794e2aed010fe64d5adda54f3d9`.

The source shape text is edited through public `set_shape_text`, serialized and
reopened while preserving geometry, properties, formatting and body text. The
checked Litchi artifact contains the sentinel exactly; the checked LibreOffice
artifact contains the sentinel plus its terminal paragraph newline. Current
generator bytes are not claimed to reproduce the historical checked changed
artifact byte-for-byte.

`rust-ci` now runs the existing offline native-resave evidence verifier and
watches its OLE2/OOXML/RTF/ODF inputs, tools, logs and outputs. It verifies
checked hashes and recorded successful same-format save outputs; it never
launches LibreOffice or active content.

## Remaining work

This tranche enables honest measurement; it does not optimize the new paths.
Formatting/media-heavy semantics, malformed and security inputs, additional
real producers, changed compressed/code-page publication and broad structural
edits remain open. Any production change must start from matched before
profiles and retain the exact stream/no-op and native readback contracts above.
