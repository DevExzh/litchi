# Descriptive-statistics preflight

The runner now performs an untimed correctness preflight for every selected
case after building the harness. It checks the frozen descriptive contract
SHA-256, the retained order-statistics gate lock, and the candidate freeze
before it can enter a timing loop. Each preflight command and its combined
stdout are retained in `preflight.json` and `preflight.stdout.log` beside the
phase results.

Capture is valid only after root supplies the frozen candidate manifest and
quiet-window authorization. Once capture starts, the implementation source
closure, contract, harness, and profile inputs must remain unchanged.

Do not replace a historical command path retroactively. If an owner changes a
package or executable name, retain the earlier command and add a new receipt
with the rename explanation.
