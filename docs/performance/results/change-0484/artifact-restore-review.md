# Artifact restore and OPC foundation review

Review scope: the working-tree `crates/litchi-opc/src/source_backed/artifact_restore.rs`, its inverse-publication helpers in `source_backed/splice.rs`, the source/budget adapters in `source_backed.rs`, the expected-artifact publication API, and the replay envelope in `splice.rs`, as re-reviewed on 2026-09-09. This is a read-only review; no production file was changed.

The durable restore seam is semantically complete. It authenticates the current and explicitly supplied original artifacts by length and complete SHA-256 before touching the sink, permits independently reopened original providers, uses the current package context for operation I/O/work/memory/output when available, observes both contexts, charges output once, rehashes the original while copying, and reports exact accepted output on post-output failure. `SourceArtifact::fingerprint` has a 64 KiB memory reservation, an actual-capacity check, checked Work charges, and typed source-error mapping (`source_backed.rs:2755-2832`). The inverse fingerprint helper has the same Work and typed source mapping (`splice.rs:1700-1737` in the current tree).

`SourcePartSplicePublication::candidate_artifact_len` is populated from the accepted `HashingSink` count on both exact no-op and decoded publication paths and is exposed by the publication API (`splice.rs:614-620`, `886-893`, `1481-1485`). Immediate inverse publication checks the candidate length before hashing (`splice.rs:1495-1558`).

The consuming `write_to_stream_with_expected_artifact` API previews a private shallow plan into `io::sink()` with the same package context, Work/Memory/source checks, and no `OutputBytes` charge (`splice.rs:510-560`, `563-576`). A preview mismatch or failure leaves the caller sink untouched and unwraps internal `IncompleteOutput`; actual emission charges output once and reports its exact accepted prefix when the expected artifact changes after output begins. Generic exact snapshot copies now reserve and capacity-check their 64 KiB buffer in both accounting and non-accounting paths (`source_backed.rs:10759-10913`), while no-op preview copies use the same checked reservation (`splice.rs:897-991`).

The replay adapter uses the smaller of the 64 KiB default and the configured authored window (`splice.rs:1645-1649`). `SpliceAuditReader` receives that actual size, checks reported capacity before resize/provider opening, and uses it for source and authored chunks (`splice.rs:2117-2173`, `2271-2335`). The 512-byte integration case crosses multiple source and payload windows for both Store and Deflate archives (`tests/source_part_splice_replay.rs:704-756`).

## Validation receipts

The retained commands, stdout/stderr, source manifests, and receipts are under [validation/](validation/) and [validation-sources/](validation-sources/).

| Receipt | Result |
| --- | --- |
| `opc-all-tests-dev37` | Exit 0; 571 passed, 1 ignored; all features and doctests |
| `opc-no-default-tests-dev38` | Exit 0; 549 passed, 1 ignored; no default features and doctests |
| `opc-clippy-dev39` | Exit 0; all features and all targets with `-D warnings` |
| `opc-rustdoc-dev40` | Exit 0; all features, no dependency docs, `RUSTDOCFLAGS=-Dwarnings` |
| `opc-format-dev43` | Exit 0; rustfmt `--check` on the OPC source and focused tests |

All five current receipts report `source_unchanged: true`. The formatting receipt is after dev37–dev40; the intervening OPC source changes are whitespace-only (two `source_backed.rs` `map_err` indentation hunks and one focused-test line wrap). DOCX development continued independently. The focused suites contain 19 replay-provider tests, 13 artifact-restore tests, and 29 fixed-splice tests. They cover expected-artifact preview acceptance and zero-output mismatch/failure, one actual output charge, exact-prefix emission failures, fresh providers on every pass, 512-byte multi-window replay, replay truncation/extra/hash/provider errors, source changes and cancellation, independent artifact reopening, distinct/shared contexts, and memory refusal/release.

No remaining concrete OPC blocker was found in the current source. The prior replay-window allocation gap is closed by the capacity check in `SpliceAuditReader::new`; dev37 and dev38 exercise the guard's normal exact-capacity path and the new preview/window paths after that change. No custom allocator rejection test is claimed. The broader DOCX stream integration, durable forward application, native-consumer/fuzz checks, scaling measurements, and final source/binary custody remain outside this OPC review.

## Positive checks

- `artifact_restore.rs:48-194` checks both lengths and both complete fingerprints before constructing the output sink; the original is authenticated again in `copy_authenticated` while bytes are emitted.
- `check_inverse_state` checks current and retained source freshness plus both context authorities, and `finish_inverse_publication` preserves source-change precedence with the exact accepted count (`splice.rs:1655-1698`).
- `write_to_stream_with_expected_artifact` runs its private preview sequentially and only the real publication constructs the output-budgeted sink. Preview progress cannot be reported as caller output.
- `read_source_at_with_context` rejects provider overreports; replay proof length is checked against finite fragment/candidate limits before plan retention, and replay EOF/digest are authenticated before acceptance.
- Every newly reserved copy/replay workspace is released by scope on success and failure, while output reservations commit only accepted sink bytes.
