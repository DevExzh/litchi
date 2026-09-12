# Baseline test review: source-backed exact no-ops

The two managed exact-no-op cell-value tests had stale memory bounds. The
scalar fixture observed 66,196 bytes against a 660-byte limit, and the
multi-sheet fixture observed 66,535 bytes against a 999-byte limit. The two
row-visibility baseline failures had the same shape: 65,953 versus 417 bytes
and 65,930 versus 394 bytes. Each excess is the fixed 65,536-byte publication
copy window, rather than detached source payload ownership.

The OPC owner makes this explicit: `SourceBackedPackage::write_exact_source`
delegates to `write_exact_snapshot` at
`crates/litchi-opc/src/source_backed.rs:9390-9392`, and managed exact-copy
publication reserves `SOURCE_PUBLICATION_CHUNK_BYTES` as
`Resource::Memory` at `crates/litchi-opc/src/source_backed.rs:10961-10980`.
Selected worksheet payloads remain separately retained by the managed cache;
output bytes use the separate output budget.

The repair adds a named 64 KiB scratch allowance to the four exact-no-op
fixtures. It also gives each fixture a deterministic large pseudo-random
untouched `/xl/media/unused.bin` member (256 KiB for cell values, 128 KiB for
row visibility), then asserts that both the untouched member and complete
source artifact exceed the available memory allowance. Existing byte-for-byte
publication checks and post-drop `Resource::Memory == 0` checks remain in
place, so the tests continue to prove source sharing without permitting whole
archive detachment. Signed changed-publication and protected-sheet refusal
assertions are unchanged.

Final SHA-256 identities of the repaired test files:

- `crates/litchi-xlsx/tests/source_backed_cell_values.rs`: `7ae1f26a2965ee0d6ccb67617b288b232de5668dc9bd51b26e0b7e8684faa00a`
- `crates/litchi-xlsx/tests/source_backed_row_visibility.rs`: `2a1b9631a6f4e642bf29e096193ef12feff6c9ae9389f4a1d2a40da574600f32`

The unchanged-production baseline receipt reports all-features tests at
`1263 passed, 0 failed, 0 ignored, 0 filtered`; formatting, Clippy, rustdoc,
owner checks, and boundary checks also passed. The candidate fifth full-suite
receipt reports `1284 passed`.
