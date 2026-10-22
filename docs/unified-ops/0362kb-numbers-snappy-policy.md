# 0362kb — Numbers retained Snappy policy

- Contract: [0351 §2.5](0351-api-contract.md#25-configured-limits-items-18-19).
- Base: `b34cb1b1049b26e183a09f8184bb0f691e917398`.
- Status: source and consumer migration complete; validation is recorded below;
  native timing is not part of this owner-policy slice.

Every production Numbers owner that encodes an IWA member now derives one
`SnappyLimits` value from the retained physical/catalog policy and passes it to
`SnappyStream::compress` and `maximum_compressed_len`. This includes creation,
metadata, names, sheet order, structure, appearance, locking, relocation,
sorting, table dimensions/headers/titles, comments and replies, chart
arrangement, image adjustments, cell controls and all nine cell-format owners.
The native pop-up graph receives the same profile instead of constructing a
default one. Test-only fixture builders in `package.rs` and `package/extractor.rs`
retain their defaults because they do not represent opened owners.

The profile is used for both admission and encoding. `maximum_compressed_len`
remains the conservative encoded-output bound; the Snappy encoder may use its
reserved output vector as scratch when that bound fits, and otherwise uses its
bounded temporary frame path. No caller substitutes the decoded input size for
the output bound. Direct popup, cell-format and comment boundaries retain core
Snappy limit and allocation failures as typed family errors, including the
comment family's distinct `CompressedBytes` axis.
Chart-data and chart-metadata projections carry that axis through their
arrangement-error mappers instead of flattening it to wire bytes.

The owner boundary tests exercise both a popup replacement and a direct comment
reply. They compute the admitted member limit from the actual target output,
accept the exact limit, and reject one byte below it with a typed
`CompressedBytes` error while leaving the source unchanged. The first version
of this test exposed a separate accounting bug: popup `TransactionBudget`
treated the per-entry ceiling as a cumulative ledger across the source and
candidate members. The corrected budget checks each actual entry against
`EntryBytes`; `TotalEntryBytes` remains the aggregate ledger. The retained
physical and state profiles were independently checked and preserved the
explicit caller values, so the refusal was not a lost-policy or defaulting
failure. Both focused tests pass with exact output and source-preservation
assertions.

The changed source has no common codec or `PROGRESS` edits. The focused mapping
test checks compressed-stream limit and allocation classification. Validation:

```text
cargo fmt --all
cargo check --offline -p litchi-numbers --all-targets --all-features
cargo test --offline -p litchi-numbers --all-targets --all-features
cargo clippy --offline -p litchi-numbers --all-targets --all-features -- -D warnings
```

The capped Numbers checks and workspace formatting must pass before this record
is accepted. The focused boundary tests are complete; the full check, test and
Clippy commands remain to be rerun after the final typed-mapping change. No
native or profiler result is claimed here; the W5 owner timing work remains
separately recorded.
