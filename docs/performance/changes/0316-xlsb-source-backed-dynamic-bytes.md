# Change 0316: XLSB source-backed dynamic byte ingress

## Scope

`litchi::sheet::open_xlsb_workbook_from_bytes_dyn` and its explicit
`_with_limits` variant now use the XLSB source-backed owner and facade adapter.
The explicitly typed `open_xlsb_workbook_from_bytes` APIs remain eager and
unchanged.

## Deferred semantics

The dynamic byte entry point makes one fallible, exact ownership copy of the
caller-provided bytes into an `Arc`-owned positional source. It validates the
OPC catalog and XLSB workbook graph at open, but defers worksheet payload
extraction and semantic parsing until a worksheet is selected. The adapter
retains the same source and read limits so existing unsupported-feature
compatibility fallback behavior remains available.

The source-backed owner preserves the complete XLSB catalog. Consequently,
active catalog positions are mapped to logical worksheet ordinals without
discarding chart, dialog, or macro tabs, and active worksheet capability errors
retain their existing typed behavior.

## Resource policy

The input length is checked against `ReadLimits::max_input_bytes()` before the
ownership allocation. The copy uses `try_reserve_exact`, and OPC/source-backed
construction receives the caller's exact limits. No generic format detector or
eager XLSB constructor fallback is used by this dynamic byte ingress.

Focused facade coverage exercises deferred malformed worksheet reads, source
ownership after caller mutation, mixed-tab active mapping, XLSX rejection, and
the exact input-byte and ordinary-part limit boundaries. Arbitrary non-package
bytes are also rejected by the XLSB-only entry point.

This change makes no RSS, OOM, throughput, latency, or allocation-performance
claim. Such claims require an independent benchmark and memory-profile gate.

## Validation

Validation ran serially with `CARGO_BUILD_JOBS=1` in one isolated target
directory:

- XLSB-only facade library tests: 45 passed.
- XLSB facade integration tests: 23 passed.
- Strict XLSB-only Clippy with `-D warnings`: passed.
- XLSB-only rustdoc with `RUSTDOCFLAGS="-D warnings"`: passed.
- `rustfmt` and `git diff --check` for the changed Rust sources: passed.
