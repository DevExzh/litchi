# 0406 — shared IWA publication admission substrate

- Status: implementation slice on `uo2/iwa-core-publication-admission`, cut
  from coordinator source `7606f8315d05fba42086ba93ea316918cdb3aa07`.
- Scope: add prepared, source-bound serialization admission for the shared
  Archive and IWA Snappy layers. No Pages, Numbers, Keynote, unified-ops,
  ledger, native queue, fixture, or PROGRESS changes are included.
- `performance_claim`: none. Full facade and full-File qualification remains
  required after each format owner adopts the shared route.

## Public substrate

`Archive::prepare_encoding(limits, admit_work)` performs allocation-free framing, metadata,
retained-header, and aggregate-limit checks and returns
`ArchiveEncodingPlan`. `ArchiveEncodingPlan::requirements()` exposes source-
derived output, payload/header, work, scratch, retained, identifier-index,
message/object, and allocation-event quantities. `execute()` reserves the exact
identifier index and output, performs duplicate-ID validation, projects headers,
and emits bytes without falling back to the legacy host route. The immutable
archive borrow prevents a prepared plan from becoming stale. Legacy
`Archive::to_bytes` and `Archive::encoded_len` remain unchanged.

`SnappyStream::prepare_compression(input_len, limits, admit_work)` returns
`SnappyCompressionPlan`. The callback is invoked in one phase unit before each
bounded preparation traversal; its failures are terminal. Its requirements
include output capacity, frame count, possible fallback frame capacity, the real `snap 1.1.2` encoder table
workspace, per-frame table-reset work, and bounded allocation events.
`execute(data)` checks the admitted input length and uses the existing framing
algorithm. The fallback capacity is explicitly an upper bound because the
selected path depends on prior compressed frame lengths; empty input reports no
heap output or encoder allocation.

The encoder workspace figures follow the vendored dependency source used by
the locked build (`snap 1.1.2`, `raw::Encoder::block_table`): a block needing
more than the 1,024-entry stack table allocates one reusable 16,384-entry
`u16` table (32 KiB), while each selected table slice is reset per frame.

## Correctness boundary

Focused tests compare prepared output byte-for-byte with legacy output for a
33-object archive, full metadata, retained unknown ArchiveInfo fields, and a
tight Snappy compressed-stream ceiling. They cover duplicate identifiers,
object limits, invalid zero limits, exact input-length refusal, and plan output
readback. No native application acceptance is claimed. The shared path still
requires format-owner adoption and final host-vs-leaf/full-File requalification.

## Verification

- `cargo fmt --all` — pass.
- `cargo test -p litchi-iwa-core --lib --offline` — pass (95 tests).
- `cargo clippy -p litchi-iwa-core --lib --tests --offline -- -D warnings` — pass.
- `cargo doc -p litchi-iwa-core --no-deps --offline` — pass.
- `python3 tools/check_api_glossary.py --profile iwa-core` — pass.
- `python3 tools/check_crate_boundaries.py` — pass.
- `python3 tools/iwa_exit_ledger.py --check` — pass.
- `python3 -m unittest tools.test_no_3rdparty_references` — pass (4 tests).
- Staged repository size and binary guards — required before commit; no
  performance evidence or binaries are added.

The core slice does not alter facade budgets. Format owners must charge
preparation callback units during planning, then charge the returned execution
requirements once before calling `execute`, preserve their own candidate
readback/unknown-content checks, and rerun their direct host-vs-leaf gates on
the final integrated source.
