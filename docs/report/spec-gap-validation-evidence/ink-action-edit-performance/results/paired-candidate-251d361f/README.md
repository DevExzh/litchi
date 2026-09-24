# Approved paired candidate receipts

This directory retains the raw receipts for the approved candidate arm of the
matched InkAction capture. The comparison is candidate-minus-baseline within
each bounded lane; it makes no package-wide performance claim.

- Source pin: `ab94a4d7a02053765bf4c70b4af5022273b821e5`.
- Clean control head: `251d361f951a0cfc24fb57e0e0ad0f76044143a0` (`ab94 + 03def`).
- Arm: `candidate`; guard: `03defc9d46a2f8ad2d9e5e0943586dec242e96a6`.
- Reviewed source change: `c8d60d5632d79351c325e3e6f1eace2d901b312c`.
- Toolchain: Rust/Cargo 1.95.0; `--release --locked --offline`; three fresh
  processes, two warm-ups, and twenty measured samples for each of 34 lanes.
- Binary digest: `692599ed0184650ccab25b1ef2283d2763a2fcf839023126bda61bcc060c1f29`, retained in `binary.sha256` and `binary-after.sha256`.
- Source manifest: `4f10008d37c28a2296deab34655939a565fe17b219d14003efba8e0f3261c3f6` before and after the run.

The profile was run with `PROFILE_ARM=candidate PROFILE_FROZEN=1
RUSTUP_TOOLCHAIN=1.95.0 WARMUP=2 SAMPLES=20 PROCESSES=3`. Its external
checkout and Cargo target were removed after verification; JSON allocator
receipts, `/usr/bin/time -v` RSS/timing receipts, stderr receipts, manifests,
commands, build provenance, and the recomputed report remain here.

`verification.json` records the passing 34-lane validation. See the [matched
comparison](../matched-report.md) and the [pair provenance](../paired-controls-provenance.txt).
