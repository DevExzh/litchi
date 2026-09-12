# Approved paired baseline receipts

This directory retains the raw receipts for the approved baseline arm of the
matched InkAction capture. The comparison is candidate-minus-baseline within
each bounded lane; it makes no package-wide performance claim.

- Source pin: `f1cb119361af9ea2227d27050e41915a9a92ae04`.
- Clean control head: `e5c18ca02e120ea1d91e495d4cf58df6d66573be` (`f1cb + 03def`).
- Arm: `baseline`; guard: `03defc9d46a2f8ad2d9e5e0943586dec242e96a6`.
- Toolchain: Rust/Cargo 1.95.0; `--release --locked --offline`; three fresh
  processes, two warm-ups, and twenty measured samples for each of 34 lanes.
- Binary digest: `0c4657b688e66ef55ae2d0ec61d06bde235d26f4c602a90d80d314686ef46476`, retained in `binary.sha256` and `binary-after.sha256`.
- Source manifest: `eb61be9257ff6035a61089a7b77fb581595e00e95b38e17d5a32cf19be92db8d` before and after the run.

The profile was run with `PROFILE_ARM=baseline PROFILE_FROZEN=1
RUSTUP_TOOLCHAIN=1.95.0 WARMUP=2 SAMPLES=20 PROCESSES=3`. Its external
checkout and Cargo target were removed after verification; JSON allocator
receipts, `/usr/bin/time -v` RSS/timing receipts, stderr receipts, manifests,
commands, build provenance, and the recomputed report remain here.

`verification.json` records the passing 34-lane validation. See the [matched
comparison](../matched-report.md) and the [pair provenance](../paired-controls-provenance.txt).
