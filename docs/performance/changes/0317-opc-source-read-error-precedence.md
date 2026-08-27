# Change 0317: OPC source-read error precedence

## Scope

Source-backed OPC cold loads now apply their post-read source-freshness and
execution-context fences even when the underlying ZIP member read fails. This
closes an error-precedence gap without changing successful reads, cache
retention, or eager `OpcPackage` behavior.

## Error contract

If an archive read races with more important package state, the observable
precedence is:

1. Source-version failure, including `OpcError::SourceChanged`.
2. Execution-context failure, including `OpcError::Cancelled`.
3. The original mapped ZIP member error.

The lower-level ZIP error is preserved when both post-read fences pass. The
same rule applies to ordinary and operation-accounted part reads.

## Cache safety

All errors remain inside the existing cold-load result path. Coordinated loads
therefore publish flight failure before removing the flight, wake current
waiters, release managed reservations, and retain no payload. Allocation-bypass
loads use the corresponding failure accounting and retain no payload.

Focused regressions combine a corrupt member with an in-read source mutation
and with in-read cancellation. They assert the higher-priority typed error,
zero remaining in-flight loads, no retained entry, and released managed memory.

This change makes no throughput, latency, allocation, RSS, or OOM claim.

## Validation

Validation ran serially with `CARGO_BUILD_JOBS=1` in one isolated target
directory:

- `litchi-opc` library tests: 257 passed.
- Strict `litchi-opc` Clippy with `-D warnings`: passed.
- `litchi-opc` rustdoc with `RUSTDOCFLAGS="-D warnings"`: passed.
- `rustfmt` and `git diff --check` for the changed sources: passed.
