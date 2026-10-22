# 0361fa — Bound field-path constructor admission without a scan

- Base: `52aec31f8c`, following the unaccepted 0361f constructor candidate.
- Status: correctness validated; fresh performance qualification required.

`FieldPath::new` previously scanned every component to determine its exact
wire length even when the configured ceiling was much larger. A `u32`
component occupies at most five varint bytes. The constructor now checks
that conservative size, including the actual length-prefix width, in constant
time. If the bound does not fit, it computes the exact length before deciding.
The bound therefore neither rejects a valid tight-budget path nor fabricates
an observed size on failure. Arithmetic overflow also takes the exact path.

`FieldInfo::new` applies the same proof to its nested path. Its optional fields
are absent by construction, so admission counts that known layout directly
instead of walking the general populated metadata model. The existing limits
validation, memory/field/nesting/item checks and owned buffers remain intact.
General archive validation still checks populated mutable metadata.

The complete core suite passes: 76 unit tests, six clone tests and 31 Snappy
framing tests. All-target/all-feature Clippy passes with warnings denied.
New boundary coverage varies component values and packed lengths across
varint-width transitions, accepts exact byte ceilings and verifies both the
observed and maximum counts at one byte under. The initial compile exposed an
obsolete private helper, which was removed; that failed attempt is retained.
[Validation receipts](results/0361fa/validation.json) cite all external logs
with their sizes and hashes.

This source is a repair candidate, not accepted constructor performance.
The earlier allocation evidence and two interrupted native attempts remain
attached to their original 0361f source. The fresh `matched-v4` build compares
`c2972b860c` with the original `ddf2ac8ae2` reference under identical release
settings and resolved dependencies. All four binary preflights reproduce the
six exact fixture/output files. Four allocation legs per route agree; every
requested/peak/retained count matches the preceding candidate and stays at or
below the reference. [The build and allocation receipt](results/0361fa/evidence.json)
retains source, binary, dependency and command identities. Native latency
remains unqualified; no timing from the allocation processes is accepted.

## Interrupted v4 native attempts

The coordinator held all three Litchi workers before running the protected v4
binaries on CPU 29. `native-a1` exited 1 before its first calibration sample
because an unrelated `peculiar` Cargo/rustc release build was active. After
those processes ended, `native-a2` also exited 1 during field-path calibration
when the same external build started Chromium/`peculiar_shell` linking. The
guard retained the calibration output and before/during/after host snapshots.
Neither attempt produced a qualified block, and neither is latency evidence.
Both runners have ended; their original manifests still say `running` because
the guard interrupted them before finalization, and those raw files are kept
unchanged. The refreshed evidence inventory records both runner logs and all
artifacts by path, size and SHA-256. Worker holds were released after the
second interruption so engineering and bounded validation could continue.
