# 0362ul165uc — RTF canonical definitions and consumer migration

Date: 2026-10-26. This leaf is cut from the coordinator's current RTF gate
tip `abe9962bc33d3aa49a8ebef50585d224225208d0` in the isolated worktree
`uo2/rtf-api-glossary`. It repairs the remaining row-46 API-glossary
findings from 0362ul49 without changing the glossary policy.

## Defect and repair

The public facade already intended the canonical paths
`litchi_rtf::{Error, Result}` and `litchi_rtf::write::{Writer, Options}`, but
the implementation still defined and imported `RtfError`, `RtfResult`,
`RtfWriter`, and `WriterOptions`, then exposed the canonical paths through
aliases. That left the format-prefixed spellings in the API surface and in
the implementation graph inspected by the row-46 check.

The codec now defines `Error` and `Result` directly. The writer now defines
`Writer` and `Options` directly. The root facade re-exports those definitions
without `as` aliases, and every RTF parser, lexer, model, metadata, drawing,
review, text, edit, transport, writer, test, and facade consumer uses the
canonical names or the canonical crate path. No behavior, error variant,
writer option, parser limit, output route, or capability was changed.

## Checks

Source-only checks on this leaf passed:

* `rustfmt --edition 2024 --check` on every changed Rust source file;
* `git diff --check`;
* a source scan found no `RtfError`, `RtfResult`, `RtfWriter`, or
  `WriterOptions` identifier under `crates/litchi-rtf/src`;
* direct-definition/export inspection confirms `codec::error::{Error, Result}`
  and `codec::writer::{Writer, Options}` are used directly by the public
  facade;
* `tools/api_glossary.json` was left unchanged, and its JSON remains the
  source of the existing ban rather than being weakened.

Cargo, rustc, native producers, and the heavy gate were not run in this
source-only leaf. The coordinator should run the RTF library/tests, the
RTF rustdoc/API-glossary profile, operations RTF consumers, and the relevant
clippy rows from the recorded diagnostic attempt
`.coord/final-gate-current-remaining20-diagnostic-20261026-v3-state/46-api-glossary-attempt-current-remaining20-full53-46.log`.
