# 0362ul135f — format the active PPT correctness adapter

The final changed-Rust formatting audit found one import-order difference in
`tools/native_qualification/adapters/ppt-scoped_patch_correctness-f2.rs`.
Root applies rustfmt and migrates the policy, both v2 manifests and the exact
identity regression in the same commit. The adapter remains 3,561 bytes; its
new SHA-256 is
`3bb1deaa67f64e8cb70f7dc2cf05ad41803fa8edd2f73b492991555d260579ab`.
Historical 135e evidence retains its original identity.

The change affects import ordering only. Forward replay, inverse replay,
serialization, workload and baseline source provenance are unchanged.
Standalone adapters are outside the root Cargo workspace, so the final gate
retains a changed-path rustfmt supplement in addition to workspace formatting.

Validation: rustfmt checks for both active Rust adapters, `git diff --check`,
and all 33 current-native/no-3rdparty Python tests pass. No Cargo, native,
metadata, timing or allocation workload was run.
