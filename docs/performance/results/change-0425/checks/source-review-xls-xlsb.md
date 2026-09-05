# 0425 XLS/XLSB constant-chunk source review

This is a read-only source review of the XLS and XLSB constant-chunk migration
against baseline `340cc91ae2bdec338dfe7682b4b5d8c219a2d288`. No Cargo command,
test command, benchmark, profiler, or CPU workload was run for this review.
The change is compiler-directed iterator maintenance; this review records no
performance claim.

## Scope and iterator semantics

The baseline had 102 `chunks_exact`/`chunks_exact_mut` occurrences in
`crates/litchi-xls` and `crates/litchi-xlsb`. The final tree migrates the 101
constant-width occurrences to `as_chunks::<N>()` or `as_chunks_mut::<N>()` and
keeps the one runtime-width loop in
`crates/litchi-xls/src/pivot_table/codec.rs` as
`data.chunks_exact(line_size)`. `line_size` is derived from the parsed
PivotTable payload, so it must remain a runtime iterator rather than become a
const generic.

For constant widths, each migrated loop selects the tuple's full-chunk slice
and iterates it in source order. Mutable sites select the mutable full-chunk
slice and retain their original mutation order. The migrated widths are
nonzero constants, including the 2-, 4-, 6-, 8-, 12-, 18-byte and named fixed
record widths used by these crates; the `as_chunks` precondition therefore
does not introduce a new runtime failure path. The API is available with the
workspace's Rust 1.89 minimum version because `as_chunks` stabilized before
that version.

The three consumers that depended on a `chunks_exact` remainder preserve that
check explicitly:

- `crates/litchi-xls/src/data_validation/codec.rs` destructures
  `characters.as_chunks::<2>()` and rejects a nonempty remainder after decoding
  the complete UTF-16 units.
- `crates/litchi-xls/src/validation.rs` keeps the remainder in the wide-name
  decoder and in `valid_utf16le`; both retain the prior odd-byte rejection
  behavior.
- `crates/litchi-xls/src/validation.rs` retains the declared name length before
  decoding the wide name and keeps the remainder in the validation result, so
  odd-byte names are still rejected after the same bounded slice.

The other migrated sites either have an existing exact-length/even-length
precondition or intentionally ignored the old iterator remainder. No site
silently drops a remainder that was previously inspected.

## Array conversion and error ordering

`as_chunks` yields references to fixed-size arrays. The direct
`u16::from_le_bytes(*chunk)` and `u32::from_le_bytes(*chunk)` forms are
equivalent to the former `TryFrom` conversions at sites where the preceding
length check already proves a complete chunk. In particular:

- `crates/litchi-xls/src/differential_format/codec.rs` checks even length before
  decoding UTF-16 units; its invalid-surrogate result remains the subsequent
  error.
- `crates/litchi-xlsb/src/host/web_extension_bindings.rs` checks the complete
  encoded record length before decoding its UTF-16 units; the old conversion
  error was unreachable under that check, while `String::from_utf16` retains
  the meaningful malformed-unit error.
- `crates/litchi-xls/src/list_object/codec/semantic/parse/source.rs` slices the
  exact count-derived byte range before decoding four-byte IDs, preserving the
  existing truncation error before iteration.
- `crates/litchi-xls/src/records.rs` compares the copied two-byte array with
  `[0, 0]`, retaining the same UTF-16 terminator position and resulting byte
  offset.

The existing count, size, overflow, cursor, and alignment checks remain before
the corresponding iteration. UTF-16 decoding and record-specific validation
therefore retain their prior order relative to malformed input errors. The
migration does not alter callback, publication, or source-version behavior;
it changes only the representation of fixed-width iteration.

These findings are bounded by the accepted correctness and validation
constraints in ADR 0005, ADR 0006, ADR 0008, ADR 0012, and ADR 0016. No source
correctness blocker was found in the final XLS/XLSB tree. Build and test
verification remains coordinator-owned and is deliberately not claimed here.
