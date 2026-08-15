# Change 0135: native XLS fixed-width numeric publication evidence

Date: 2026-08-15

## Scope

This change adds four opt-in `tools/perf-baseline` selectors:

- `xls_numeric_eager_number_edit_save`
- `xls_numeric_source_backed_number_edit_save`
- `xls_numeric_eager_rk_mulrk_edit_save`
- `xls_numeric_source_backed_rk_mulrk_edit_save`

The Number pair reuses the deterministic comments CFB and edits
`Untouched!E21` from `42` to `43`. The packed pair uses a separate deterministic
native XLS CFB containing one standalone RK record and one two-cell MulRK
record; one transaction edits the standalone cell and both packed cells with
exactly representable replacements. The corpus carries opaque sibling streams
and metadata so publication must preserve complete topology and untouched
member bytes.

## Timing and evidence contract

Corpus construction, source ingress, expected eager/source-backed outputs, sink
capacity reservation, no-op/fingerprint checks, patch apply/inverse/stale
checks, unsupported and security refusals, complete `Snapshot`/`Workbook`
reopen/readback, and the real-producer `54016.xls` reopen/inverse gate are
outside the timer. Each measured sample separately records transaction
creation, `set_number` or `set_numeric`, eager `commit` versus
`commit_source_backed`, and publication. Both paths publish complete target
bytes to equivalently configured preallocated bounded `CountingSink`s (64 KiB
maximum write); `elapsed_ns` is
the arithmetic sum of the four separately timed phase vectors, not a continuous
wall-clock timer.

Generic `source.read_calls`/`source.read_bytes` carry the owned source-ingress
counters. `source.xls_numeric` reports input/output CFB and Workbook sizes, complete
target materialized bytes on both implementations, source-backed splice /
replacement / changed-span / source-target fingerprint vectors, source ingress
scope, sink bytes/write calls/digests, and the explicit owned-input scope.
The source-backed commit retains a complete reopened target snapshot, so this
change does not claim bounded artifact memory or file-positional ownership. It
also makes no allocation/RSS, physical-I/O, speedup, or broad real-producer
claim.

## Correctness gates

The runner requires deterministic per-case output digests, source-backed equal
Workbook lengths, unchanged CFB stream/storage topology and opaque payloads,
storage-family-preserving numeric readback, exact forward/inverse patches and
stale-source refusal. It checks exact no-op identity and fingerprints, atomic
unsupported edits, and signed, macro, protected, and encrypted/refused paths.
The 54016.xls producer fixture is used only for the untimed RK/MulRK reopen and
inverse gate.

The selectable matrix is now 261 names while the historical default remains 36
cases / 198 records. This is correctness and CRUD-coverage evidence only; no
release ABBA result is accepted by this change.

## Verification

Focused one-sample smoke:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 1 \
  --case xls_numeric_eager_number_edit_save,xls_numeric_source_backed_number_edit_save,\
xls_numeric_eager_rk_mulrk_edit_save,xls_numeric_source_backed_rk_mulrk_edit_save \
  --json target/perf/xls-numeric-0135-smoke.json
```

The harness source was checked with strict `-D warnings -D deprecated` Clippy,
formatting, and a clean diff review. No production or test source outside the
assigned harness/docs paths is changed by this note.
