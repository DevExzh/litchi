# Provisional isolated XLSX SVG lifecycle gates

These receipts cover the exact source hashes in `inputs.json`, over committed
base `dc8087e81`. All 1,105 library tests, 37 lifecycle tests, and 14 SVG read
tests pass. Strict Clippy and warning-denied rustdoc pass. This is correctness
and build evidence, not production approval or a performance result. Independent
resource review remains pending, including the mixed repeated-selector fallback
and complete preallocation admission.

The isolated checkout excludes unrelated OPC, DOCX, and InkAction working-tree
changes. `source-overlay.tar.gz` retains all 13 copied inputs, including the
pinned Cargo.lock and native read fixture. Reproduce by checking out the base,
extracting that archive at the checkout root, and running the commands in
`receipt.json`. Hash every extracted file against `inputs.json` first.

The first focused invocation failed during setup because the ignored root
Cargo.lock was absent. Its log is retained separately; the rerun used the
recorded lockfile with `--locked --offline` and passed. No source drift was
observed between the shared tree, isolated copies, and the captured inputs.
