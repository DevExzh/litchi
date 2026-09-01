# Change 0361: bounded streaming x14ac observers

**Date:** 2026-09-02
**Status:** Implemented
**Performance claim:** none

## Decision

Implement the bounded streaming x14ac raw and active observer foundation in
`litchi-xlsx`. The MCE raw observer sees ordinary and alias duplicates before
generic duplicate validation, which preserves the byte-compatibility
information required by x14ac handling. The MCE and `AlternateContent` x14ac
byte-compatibility branch now streams; the plain fast path is unchanged.

## Processing contract

After a semantic `NonConformant` or `MustUnderstand` result, the raw-only path
can perform one-pass recovery. If that recovery later encounters an XML,
input, or limit failure, the later failure is primary while the typed prior
semantic error is retained. The raw and active observer surfaces remain
bounded by the existing stream limits.

Input uses a fixed 8 KiB `InterruptedRetryReader` with a maximum of eight
interrupted-read retries. The x14ac branch consumes the MCE stream without
materializing a rewritten XML buffer, while the unchanged plain fast path
retains its existing behavior.

## Resource boundary

The fixed 8 KiB statement covers the input reader buffer only. quick-XML
parser state, decoded values, observer allocations, and collection overhead
are outside that fixed-buffer claim. With x14ac `capture_rows=true`, a
`BTreeMap` may retain rows up to the configured `ROWS` limit. This is bounded
streaming integration, not selected-cell or full-worksheet streaming, and it
establishes no latency, RSS, or OOM-safety claim.

## Validation

Focused and crate-scoped validation passed:

- MCE recovery tests: `7/7`;
- raw-attribute tests: `4/4`;
- x14ac focused tests: `12/12`;
- worksheet tests: `35/35`;
- `litchi-ooxml-common` library tests: `234/234`;
- `litchi-xlsx` library tests: `813/813`.

These are correctness and bounded-streaming integration observations only;
`performance_claim: none`.

## Residual scope

Selected-cell and full-worksheet streaming remain later work. The x14ac
`capture_rows=true` mode can retain a `BTreeMap` up to configured `ROWS`, and
quick-XML and observer allocations are not covered by the fixed 8 KiB input
buffer. No latency, RSS, or OOM-safety claim follows.
