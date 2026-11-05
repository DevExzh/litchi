# 0362ul89f — Combined PPTX compiler return

Date: 2026-10-25. This is a required compiler repair, not runtime or
performance acceptance. Only redaction audit 22 is deferred.

The first actual two-row combined capture freezes
`2915ef176285376c4f7c996d5ed2317cf12b72f7` (tree
`06a827d34a5e549d57aa2c277e3f6f12ea23d810`). Formatting passes for all 384
Rust paths changed from `ad138ffcedb5afc8ddbdbcc69c1e0c27e064cc77`.
`cargo check --offline --locked -p litchi-pptx --all-targets --all-features`
then exits 101 in the OPC dependency, reporting three unnecessary mutable
bindings in `validate_combined_relationship_overlay_limits`.

The repair removes `mut` from the three immutable catalog totals. The
separate adjusted totals remain mutable; their overlay calculations, checks
and observations are unchanged. Scoped rustfmt and `git diff --check` pass.
A fresh guarded compiler capture is still required.

The actual capture is retained outside Git at
`(deleted-coord-artifact)`.
Its approved driver is `run_f5.py --go`; runner SHA-256 is
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.
The coordinator verifies commands, source/lock fences, logs, target traces,
resource guards and empty terminal ownership. Final target usage is
179,929,333 bytes. No later check or test is inferred to pass.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `plan.json` | 32,636 | `f3adf2e5c1fd84c908bb98935f7fa4ea147af0df58ede38b2331b9a02bff87f0` |
| `rows.json` | 31,268 | `1572a43c2c879b8f4bf0d3bb4b6a88db21fc8d9d8e5ba4a86a013bda2e47a3d4` |
| `result.json` | 32,935 | `480cef91535930ad67e8fb2f3439d7d656a64296ac5571e49ac8a57d80f1b75d` |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 23,837 | `7f00313195f6ee5b8baa4090f8c89a0683eddfb6f32afaea738fbc5f79ae7419` |

An earlier detached driver launch (PID 2956690) exits before any row or
result artifact appears. It is retained as a launch failure, with no cause
asserted. The actual capture runs through persistent session 40633, driver
PID 2961296; the driver and all traced Cargo processes are absent afterward.

Combined tests, strict Clippy, exact transcript regeneration and performance
qualification remain open, followed by final integration and full acceptance.
