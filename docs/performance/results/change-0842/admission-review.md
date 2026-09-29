# 0842 independent pre-timing review

The independent reviewer confirmed 24 OPC qualifications with exact paired
archives and unchanged local spans, 18 codec archives with explicit Deflate
and same-leg build/source checks, and exactly two resolved feature additions:
flate2/runtime_detection and zlib-rs/std. Final quality and the fixed native
schedule remain required; this review is not an adoption decision.

The reviewer identified three issues that the coordinator addressed before
native admission:

- Write/changed perform duplicate output decoding after their clocks. The
  retained probe keeps this fixed validation-interleaved protocol; selection.md
  explicitly states the cache/thermal/frequency limitation and excludes native
  causal phase attribution. Both binaries use the same probe source.
- Codec route intent alone did not independently prove the compression method.
  codec_check_v2.py now checks every local and central method is Deflate and
  binds each codec capture to its own leg's build and source freeze.
- The full workspace all-target check hit an iWork example E0308. The same
  failure was reproduced with the original dependency declarations. The user
  excludes iWork; its source remains untouched. The final quality script checks
  workspace libraries/tests and all affected ZIP/OPC/OOXML targets, in addition
  to workspace library Clippy and affected-owner documentation/tests.
  Full workspace rustdoc also failed on two unresolved links in the excluded
  litchi-numbers-wire crate; the baseline control and scoped replacement gate
  are recorded in workspace-doc-limitation.json.

Other points were resolved by existing explicit boundaries: analyzer pass is
validity rather than adoption, zero-denominator fault ratios remain undefined,
p50/RSS are positive, and the live perf-baseline lock update is required for
its --locked build rather than an unrelated production package upgrade.

A second independent audit required canonical complete lock-package comparisons
(including duplicate package names), locked tool build receipts at admission
and closure, and original-manifest hashes for the baseline rustdoc control.
The reviewer rechecked those fixes successfully. The stronger lock guard then
caught stale native-resave local dependency edges; Cargo reproduced them with
the original compression declarations. See native-resave-lock-drift.json and
lock-integration-repairs.json. The candidate differs from that normalized lock
only in flate2's crc32fast edge.
