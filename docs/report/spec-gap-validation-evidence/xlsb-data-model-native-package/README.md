# Synthetic Data Model lifecycle in a native XLSB package

This test-only batch uses `test-data/ooxml/xlsb/date.xlsb` as a native, model-free
base. Public APIs add synthetic connections and opaque Data Model bytes, then
save/reopen, edit the minimum load version, apply the inverse, and remove the
model. No-op commit/apply preserves the package's baseline serialization.
Unrelated native part payloads remain byte-identical, and full model removal
restores the exact workbook binary from the post-connections baseline.

The model payload is intentionally synthetic and opaque. This verifies integration
with an existing native package, not compatibility with an Excel-produced Data
Model payload, calculation, refresh, or a native application's acceptance of the
output. The true native Data Model producer-evidence gap remains open. No
production code or performance claim changes in this batch.

The fixture SHA-256 is
`fbb969989aabed2e057b2ddfc4cfd10ef1f6ebe9b7697c3755283515158d5b09`.
The companion verifier hashes the actual fixture and source files; the Rust
tests concentrate on the semantic lifecycle and byte-preservation assertions.
`validation.json` records the focused commands, toolchain and retained log hashes.

The tested worktree included a pre-existing import-order-only change in OPC
`phys_pkg.rs`. That production file was left untouched and unstaged.
`worktree-input.patch` retains the exact input difference for reproducing the
source manifest from this test commit; it is not a feature change.
