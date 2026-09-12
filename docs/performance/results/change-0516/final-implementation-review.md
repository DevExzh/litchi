# 0516 final emitted-output implementation review

`performance_claim: none`

`claim_authorized: false`

This is a focused implementation review of the emitted-event worksheet parser
feed. It covers only effective changed worksheets that require
`requires_store_verification`; exact no-ops, ineffective edits, metadata-only
edits, ODF, and iWork remain outside this path.

## Snapshot and verdict

The review is bound to base revision
`c5abaef129f5ae0000a925b295cf6507b7474e3c`, the after-epoch source manifest
`38007ba3313a420bf1f5823fa8aafb2232d9f3f74bc0874ab416e302263b1df6`, and
candidate patch SHA-256
`9b3cf124904ed5970a376c26a4aabd900c212d34755490f916c60f367c8f9b3f`.

**Focused verdict: approve the three previously blocked semantic fixes.**
Decoded shared-string/shared-formula exclusion, finalizer error ownership, and
namespace-resolved emitted-event feeding are implemented consistently in the
bound snapshot. The overall 0516 retention or performance decision remains
held for the independent resource-accounting review; this document does not
authorize a performance claim.

## Findings

### Decoded shared values are excluded before speculative finalization

[`codec.rs`](../../../../crates/litchi-xlsx/src/raw/worksheet/codec.rs) now
checks the parser's decoded `cell_type` and `formula_kind` immediately after
the shared `Parser::consume_event` call. This catches both literal and escaped
values such as `t="s"`/`t="&#x73;"` and
`t="shared"`/`t="shar&#x65;d"`. A shared-string cell is therefore dropped from
the provisional feed before `finish_store` can request the package-owned
shared-string callback, and a shared formula is dropped before shared-formula
translation can expand its text. The transaction then parses the exact
compacted bytes, preserving the existing x14ac/MCE/parser owner order.

The tests
`escaped_shared_formula_type_matches_literal_and_budget_fallback` and
`shared_string_owner_error_is_typed_and_called_once` cover escaped decoding,
typed owner errors, exact-byte parity, and one callback invocation. This closes
the raw-byte detector gap from the earlier candidate.

### Finalization has one owner and preserves errors

[`compact.rs`](../../../../crates/litchi-xlsx/src/raw/compact.rs) consumes the
provisional parser once and returns `parser.finish(strings).map(Some)`. It no
longer turns finalization failures into `None`, so an ordinary eligible
finalizer error remains authoritative and is not retried through a second
parser. For shared-string and shared-formula inputs the decoded exclusion
leaves `parser` absent; the transaction makes exactly one exact-parser callback
instead. The callback-error regression confirms the typed diagnostic and
single-call behavior.

The phase order remains compact output first, exact x14ac/MCE ownership and
UTF-8/parser fallback as required, then web validation, styles, requested
changes, and publication. A feed event error or budget/safety refusal only
drops provisional state and does not become a new public diagnostic.

### Namespace and emitted-event equivalence is established for the eligible path

`normalized_start` copies the source qualified name and every raw attribute,
including namespace declarations. End events use the writer-regenerated
qualified name. The feed resolves those normalized events against the source
reader's active namespace scope, which is the same scope represented by the
copied declarations in the exact compacted bytes. The dynamic safety gate
rejects unknown namespace bindings, MCE/x14ac expanded names or declarations,
invalid UTF-8, and attribute forms that the writer could serialize differently
(`attributes_raw()` apostrophes, raw quotes, or raw `<`). The post-output gate
rechecks the MCE/x14ac-free, borrowed, UTF-8 output before retaining the feed.

The differential tests cover transitional and strict aliases, nested prefix
rebindings, default-namespace reset, namespaced attributes, empty elements,
exact compacted bytes, Store values, and web bindings. In particular,
`nested_namespace_rebindings_preserve_exact_output_store_and_web_result` and
`output_reader_aliases_and_strict_namespace_bindings_match_reference` pass
against the ordinary exact parser. No namespace or emitted-event mismatch was
found in the eligible MCE-free/x14ac-free path.

### Decoder and extension ownership

The current workspace uses the same quick-xml decoder for the source events and
the normalized feed, and the ordinary parser path remains the authority for
declared extension content. No decoder heuristic was added. The input gate
rejects common MCE/x14ac candidates before constructing the feed; the dynamic
and output gates reject any extension namespace that appears after writing.
Thus no x14ac value is accepted by the provisional parser without its
extension map having been initialized, and all extension-bearing cases retain
the exact compacted-byte fallback.

## Remaining release gates

The resource-accounting reviewer owns the final decision on the candidate
ceiling and all transient/finalization terms. The four-copy feed payload
accounting in the after snapshot addresses the latest text-materialization
concern, but this semantic review does not substitute for that independent
bound review. Retention should remain disabled for any case that fails that
review; the exact compacted bytes remain the fallback authority.

The source-bound fifth-unit receipts report `970 passed` for
`cargo test --locked -p litchi-xlsx --lib` and `1,284 passed` across the
all-features suite. The after snapshot changed only the feed payload accounting
in `codec.rs`; the root control run is the evidence owner for any additional
format, clippy, formatting, rustdoc, and boundary receipts.

## Final file bindings

The following hashes are the exact current files covered by the after manifest:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/raw/compact.rs` | `0ac488dad6162f79b5d38713a730dbf6c4269f20f323505009366de7008f39d3` |
| `crates/litchi-xlsx/src/raw/compact_output_tests.rs` | `f549dc1f814042c28793640872ed55d708afc8bfb67e19e49ba38c3eae86acb6` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `0515166610208cd8b811aaff03ac880ac85f7bce689fd119f5a2fcbe6fa2f71c` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `da769ff4964042367945bc827a3c3067b718e8714696cf4bc87dc2777d26722b` |
| `crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs` | `eb98adaf7cd5a8f8d9fe4e40aa13575f2396f07f6909528e1764ffcfc072ede7` |
| `crates/litchi-xlsx/src/workbook/edit/tests/mod.rs` | `987a3f31a791efce93d3ffc5d340b5e7765591059c781e351d8aa88e1232bb1e` |
| `crates/litchi-xlsx/src/workbook/edit/tests/output_fusion.rs` | `950993b3062ece21d37139a6b141ae4f53afb329ee17ba3a29b371e5a855e947` |
| `crates/litchi-xlsx/tests/source_backed_cell_values.rs` | `7ae1f26a2965ee0d6ccb67617b288b232de5668dc9bd51b26e0b7e8684faa00a` |
| `crates/litchi-xlsx/tests/source_backed_row_visibility.rs` | `2a1b9631a6f4e642bf29e096193ef12feff6c9ae9389f4a1d2a40da574600f32` |

