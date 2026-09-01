# Change 0359: callback-scoped verified decoded readers

**Date:** 2026-09-01
**Status:** Implemented
**Performance claim:** none

## Decision

Add callback-scoped verified decoded readers to `soapberry-zip` and
`litchi-opc`. The new path is a bounded transport foundation for later
source-backed parsers, including selected-cell XLSX work. It does not change
the current XLSX worksheet parser and makes no latency, RSS, allocation,
constant-memory, or OOM-prevention claim.

## ZIP contract

`IndexedArchive` now exposes verified Store and Deflate readers through an
HRTB callback. The borrowed `BufRead` cannot escape the callback. Its decoded
window is a fixed 16 KiB buffer, interrupted physical reads have a finite retry
budget, and invalid `consume` calls become deferred typed failures rather than
panics.

After an ordinary callback return, the implementation drains the remaining
decoded stream and verifies exact decoded size, one-byte overrun, CRC, strict
payload consumption, and compressed consumption. A callback error is retained
as a typed secondary value when a transport or archive verification error is
primary. Callback errors take precedence over accounting-only failures.
Accounting distinguishes callback-accepted decoded bytes from bytes discarded
by finalization.

A callback panic unwinds without a drain or cache publication. The callback's
side effects are tentative until the method returns success.

## OPC contract

`PartView` now exposes `with_verified_decoded_reader` and its accounting
variant. The path applies part limits, declared work charging, a 16 KiB managed
memory reservation, source-version checks, execution/cancellation fences, and
typed source-change and execution-error recovery. It does not materialize
`PartData`, admit the payload to the part cache, reserve the declared payload
size as memory, or retain a cache lock across the callback.

The public `VerifiedDecodedReaderError` retains a callback error when an OPC
error wins. Error selection is source change, cancellation/execution,
archive/transport/size/CRC verification, callback, then accounting.

The 16 KiB statement covers decoder scratch only. The pre-existing strict ZIP
layout proof is archive-wide indexed state, may allocate in proportion to the
archive catalog, and is outside this fixed-scratch claim. Later end-to-end XLSX
streaming work must either account that state separately or reuse a previously
admitted proof.

## Validation

Validation was deliberately serial and crate-scoped:

- `soapberry-zip` callback-reader tests: `4/4`;
- `litchi-opc` callback-reader tests: `6/6`;
- `soapberry-zip` library tests: `319/319`;
- `litchi-opc` library tests: `277/277`;
- OPC operation-accounting integration: `13/13`;
- OPC source-backed-reader integration: `6/6`.

One initial ZIP test assertion was corrected after its fixture changed from a
size corruption to CRC corruption. One initial OPC compile exposed the ZIP to
OPC accounting boundary and an unused import; both were corrected before the
passing runs. A mistaken module-path filter matched zero tests and is excluded
from the evidence above.

All commands used one Cargo process, `CARGO_BUILD_JOBS=1`, one test thread, an
8 GiB process ceiling, disabled incremental/debug compilation, and the single
on-disk `change-0359` target. The target's observed final/peak footprint was
381 MiB. Host availability was approximately 14 GiB with 134 GiB disk free and
exhausted swap. No parallel build or OOM occurred. These are bounded validation
observations only; `performance_claim: none`.

## Residual scope

The selected-cell XLSX path still materializes a complete worksheet and parsed
store. End-to-end improvement requires a streaming MCE/x14ac event layer,
full-EOF worksheet semantic scanning, bounded shared-string/style resolution,
and measurement under the existing serial resource protocol.
