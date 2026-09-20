# ODS value inspection performance evidence

This directory owns the reproducible process harness for the sixteen pure
value-inspection and conversion functions listed in
[`baseline.json`](../baseline.json): `ERROR.TYPE`, `ISBLANK`, `ISERR`,
`ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`,
`ISTEXT`, `N`, `NA`, `NUMBERVALUE`, `TYPE`, and `VALUE`.

The matched baseline is commit `d623f3c2ecc0c837017f700174656f0e443759a5`.
The calendar-only follow-up `08bb119708` shares deterministic helpers without
changing the baseline formula behavior. The retained gate lock is the byte
profile lock SHA256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`; the
ambient workspace lock remains a separate, non-authoritative input.

The harness will retain the previous profile's borrowing resolver, counting
allocator, bounded budget observer, source and cancellation fences, exact
read accounting, and fresh child process protocol. Its new lanes cover scalar
identity and conversion work, borrowed mixed-type references, matrix lifting,
`TYPE` array metadata, known descriptor refusal before reads, projected lazy
`IF` cache coordinates, numeric/date/fraction parsing, separator transforms,
sticky cancellation, and resource failure. The matched controls remain
unchanged and do not contain the new function names.

The semantic contract, independent Python oracle, and JSON goldens are now
present. The harness has 28 matched controls and 59 candidate cases; its
preflight checks exact scalar values, per-coordinate matrix/reference results,
the 64-cell TYPE scan, N's explicit intersection and inline `[0,0]` choices,
and zero reads for both reference-list and invalid-separator refusals.
Source-freeze hashing and the final independent gate handoff remain pending;
no timing capture is included in this preparation directory.
