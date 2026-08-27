# Change 0318: Release source-backed overlay comparison handles

## Scope

Source-backed OPC overlay publication now releases temporary decoded
`PartData` comparison handles before exact-source or changed-member
publication. The cache itself is unchanged: clean entries are not purged, and
externally held handles may continue to pin their payloads.

The release occurs after exact payload comparison and, for changed XML Parts,
after both original and replacement XML audits. This preserves the malformed
but byte-identical no-op path and the existing signature-policy ordering.

## Correctness contract

Exact no-ops still copy the complete source artifact byte for byte, including
signatures, opaque members, ZIP framing, and archive comments. Changed overlays
still raw-copy every unselected source member and regenerate only the selected
replacement members. Source freshness, cancellation, cache, limit, and
partial-output behavior are unchanged.

The focused managed test uses a cache smaller than the selected Part, forcing
the comparison read through the oversized bypass path. Its sink observes zero
managed `Memory` reservation at the first exact-source publication write, the
output equals the source bytes, and the final managed memory reservation is
zero. This is handle/reservation-lifetime evidence only; it makes no RSS, OOM,
throughput, or latency claim.

## Validation

Validation ran serially with `CARGO_BUILD_JOBS=1` in one isolated target
directory:

- `litchi-opc` library tests: 258 passed.
- Strict `litchi-opc` Clippy with `-D warnings`: passed.
- `litchi-opc` rustdoc with `RUSTDOCFLAGS="-D warnings"`: passed.
- `rustfmt` and `git diff --check` for the changed sources: passed.
