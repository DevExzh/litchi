# 0362ul94f — core test-helper compilation gate

The actual compiler-first check on `6d95f193c5c852de3b4f1e9285aebb88a7edbdb9`
passes formatting on 90 changed Rust paths and clears the prior ten compiler
errors. It then reports one unused helper, `ensure_core_only`.

That helper's only call is inside the already test-only `Slot::stage`.
This commit gives the helper the same `cfg(test)` gate. The production
transaction continues using `ensure_core_only_transaction`, which returns
the typed content-free `Refused(UnsupportedProperty)`. No public capability,
consumer, limits, assertion, or algorithm changes.

Actual state:
`(deleted-coord-artifact)`.
Command:
`cargo check --keep-going --offline --locked -q -p litchi-ooxml-common --all-targets --all-features`.
Exit 101 after 36.282 seconds; target 158,632,675 bytes. Root verifies exact
commands, clean source/three-lock fences, log/trace hashes, empty preflight
builders, no live guard errors, terminal ownership and absent Cargo PIDs.

| Artifact | Bytes | SHA-256 |
|---|---:|---|
| `rows.json` | 12,088 | `c1d6e6d213a9eccd0c04467116aa2f905081d18634a6fa51d19e39ae03951a10` |
| `result.json` | 13,172 | `dbfb284d2044ea478782580bb14e04bafff7f109d0821bd399c15a134cbc5aad` |
| `01-check-common-attempt-actual-f5-01.log` | 17,241 | `2050f99a871bb66301908592c050ad895a264b1cec6a38f7118e1c14e807dc54` |

Direct formatting and repository guards pass; fresh compiler validation is
required. Custom XML94cb and the complete combined qualification remain open.
