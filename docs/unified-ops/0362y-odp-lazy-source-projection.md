# 0362y — ODP lazy source projection

This slice repairs the ODP source-backed heap residual identified by
[0358a](0358a-w6-odf-review.md) §6–7.  The tail-append owner used to retain a
second `Arc<str>` copy of `content.xml` (and the optional `styles.xml`) after
open.  On the matched 0358a source-backed probe this accounted for the
measured 593,864-byte peak increase, or 1.108 times the prior peak.

The candidate is based on `cc76bd50dea6c9d47a38b4a923f29e7867d96f3c`.  Open
still reads every validated XML part under the existing declared-size and
UTF-8/structure checks, so malformed input and oversized input are rejected at
the same boundary.  It now drops the validation strings after open and keeps
only the validated `content.xml` byte count, package metadata, tolerance facts
and source proof.  `content.xml` and `styles.xml` are loaded through
source-version checks on the first semantic read and retained in `OnceLock`
only when a caller actually needs the projection.  Shares clone the optional
`Arc` values without reintroducing eager retention.

Tail append, deferred publication and patch replay continue to use the
bounded `SourceContentInsertionPlan` and its recorded before/after proofs.
Publication still passes the same normalized validated `content.xml` to the
common raw-preservation writer, and the source-backed tests continue to cover
exact output, no-op behavior, source freshness and equal-content forward and
inverse replay.

The focused unit test
`package::source::tests::source_owner_drops_large_validation_buffer_before_semantic_read`
constructs a presentation with a 1,048,576-byte text member and checks that
the owner has no cached XML immediately after open, then checks that the cache
appears only after `content_xml()` is requested.  The integration source
reader tests also verify that the first semantic read performs the deferred
source read and subsequent cached reads do not read the member again.

## Validation

| Check | Result |
|---|---|
| `cargo check --offline -p litchi-odp --lib` | passed |
| `cargo test --offline -p litchi-odp --lib source_owner_drops_large_validation_buffer_before_semantic_read` | passed |
| `cargo test --offline -p litchi-odp --test source_backed` | 12 passed |
| `cargo test --offline -p litchi-odp --test sequential_text` | 16 passed |
| `cargo test --offline -p litchi-odp --lib package::source_append::tests` | 12 passed |
| `cargo clippy --offline -p litchi-odp --all-targets --all-features -- -D warnings` | passed |
| `cargo test --offline -p litchi-odp --all-targets --all-features` | passed: 223 unit tests and all integration targets |

The source and test snapshots used by this candidate were below the repository
evidence limit: `crates/litchi-odp/src/package/source.rs` was 39,783 bytes
(SHA-256 `e5b16bee7be793884c71d22a64272eaf6bf4bb825afa74299c5a9f0200292bdc`)
and `crates/litchi-odp/tests/source_backed.rs` was 19,840 bytes (SHA-256
`8578c324f7ae71b0dafe937c429509dd0b6b738ea721310bf23a6e70e7e0eaec`).  The
The sequential text consumer was 18,746 bytes (SHA-256
`c778fb7e6553af135f2b763191df96b8ee7b50486667089fb3eaeb699bdd0a80`). The
compact receipt is [results/0362y/validation.json](results/0362y/validation.json).

No native timing or heaptrack run was performed while the host was occupied.
The direct cache-state/counting proof and the existing 0358a paired heap
measurements establish the mechanism; a quiet native rerun remains the
integration proof for the exact peak-byte reduction.
