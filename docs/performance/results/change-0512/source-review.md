# 0512 current XLSX source review

Read-only review at `18d4d7fe1452bf43614a6efbd5feab4adb5a174f`. No production
change is admitted by this review. Historical 0466–0470 inclusive percentages
are not current operation costs. 0472 already elides owned tags for plain
cells; 0471 rejected early rewrite-buffer release for lack of measured memory
benefit. Neither should be proposed as untried work.

## Path and obligations

`workbook/edit/semantic/transaction.rs:1023` enters commit. Effective sheets
access `sheet.store()` at 1363; the source Store parser and style validation
precede lossless snapshot scanning. `raw/worksheet/edit/package.rs` builds the
snapshot Layout before rewriting spans. Changed XML is compacted at 1672,
reparsed for required grid verification, checked for web bindings and styles,
then read back against staged changes. The bounded Store handoff remains
4,096 cells/1 MiB. Save delegates to the OPC writer.

These traversals are not interchangeable: MCE preprocessing may change the
eager parser's input, while snapshot offsets refer to original source bytes.
A combined pass must preserve preprocessing and source semantic/style errors
before snapshot errors and retain unsupported content. Retaining both RawCells
and Layout simultaneously also requires operation-local peak evidence.

## Candidate ranking before the fresh profile

1. Snapshot cell address/tag processing is precise repeated work.
   `snapshot/scan.rs:952–1002` first calls `cell_address`, whose checked
   `unqualified_attribute_value` walks attributes and normalizes unqualified
   `r`. It then calls `wire::cell_tag`, which checks and normalizes attributes
   again. 0472 avoids owned Tags for plain cells but retains this second walk.
   A bounded fused or second-pass helper needs the exact ordering proof below.
2. Conditional eager parser/snapshot fusion is larger but only plausible on
   an unchanged borrowed MCE-preprocessing result. It needs memory evidence,
   complete validation equivalence and no-op overhead review.
3. Eager shared-formula resolution scans all cells even when none are shared;
   any proven no-shared fast path must retain formula/type/string/style,
   duplicate and allocation-error behavior. Do not add a monotonicity pre-scan:
   the current unstable sort already handles ordered input efficiently.
4. MCE processing is conditional on its namespace marker. Optimize it only
   when the current profile shows a paid cost; preserve namespace directives,
   AlternateContent, unsupported markup and error precedence.
5. Compaction scans Start attributes for `xml:space`, then again in
   `write_start`. This is a separate smaller candidate; byte/error differential
   coverage is required before combining scans.

## Snapshot error-order proof required for any candidate

The address helper completes checked raw syntax/duplicate iteration first,
while decoding unqualified `r` as encountered. Thus an invalid `r` entity can
precede a later malformed or duplicate attribute. Coordinate syntax, row
agreement, inferred-column limits and `last_column` mutation then precede
unknown/qualified value decoding by the tag phase. Invalid tag/attribute UTF-8
in that phase must remain after coordinate validation. Qualified `x:r` is
not an address. Preserve normalized values, order, unknown attributes,
untouched byte spans, generated coordinates and plain-cell `Option<Tag>`
elision. An ordinary-case shortcut must fall back without changing any of
these failure boundaries; no partial Layout may escape.

The governing constraints are ADRs 0001/0003/0005/0006/0011/0018: immutable
atomic publication, typed failure, bounded fallible allocation, original-source
provenance, preservation, and XLSX/OPC ownership. All 30 previously read ADR
files are hash-verified unchanged. Fresh profile results, rather than this
source ranking alone, determine the next implementation.
