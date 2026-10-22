# 0362n — DOCX/PPTX metadata commit triad

This slice migrates the eager DOCX and PPTX metadata families to the shared
commit shape. It is source work only; no performance qualification is claimed.

## Source and implementation

- Base: `933ce43de3625fad117dd94c6d11793634008dea`.
- Implementation: `2eca82b0f8bdc8701e3d2add197b434ec7b21199`.
- Changed files:
  - `crates/litchi-docx/src/metadata.rs`
  - `crates/litchi-docx/src/package/package/metadata.rs`
  - `crates/litchi-docx/tests/core_properties.rs`
  - `crates/litchi-pptx/src/metadata.rs`
- DOCX re-exports `litchi_core::edit::Diagnostics` and aliases
  `litchi_core::edit::Commit<Package, Patch>`.
- PPTX aliases
  `litchi_core::edit::Commit<Package, Patch, properties::Diagnostics>`.
- Both `Package::apply_metadata` methods return their family commit. Direct
  consumers use `into_package`, while the patch and diagnostics remain
  available through the triad accessors.
- DOCX authenticates the shared property patch before a combined core-property
  publication and does not apply an authenticated no-op a second time. PPTX
  uses the existing package duplicate fast path for an authenticated no-op,
  preserving retained policy, memo state and signed bytes.

## Validation

The following commands were run from the metadata worktree with the capped
`.coord/w8-file/env.sh` environment, `CARGO_BUILD_JOBS=2`,
`RUST_TEST_THREADS=2`, and a worktree-local target:

```text
cargo fmt --all -- --check                         pass
cargo test -p litchi-docx --lib metadata            42 passed
cargo test -p litchi-docx --test core_properties    2 passed
cargo test -p litchi-docx --test package_signature_policy  9 passed
cargo test -p litchi-pptx --lib metadata            120 passed
cargo test -p litchi-pptx --test pptx_metadata_families  6 passed
cargo test -p litchi-pptx --test package_signature_policy 10 passed
cargo check -p litchi-docx -p litchi-pptx -p litchi-ops --all-targets --all-features pass
cargo clippy -p litchi-docx -p litchi-pptx -p litchi-ops --all-targets --all-features -- -D warnings pass
```

The command output was observed in the worker terminal and was not retained as
an external log before this record was written. No binary, raw capture or
large evidence file is part of the commit. The worktree target was 1.8 GiB
and is untracked build output.

## Integration dependency

The W4 shared-properties scope/no-op repair is intentionally not included in
`2eca82b0f8`. Coordinator integration must apply the reviewed dependency
`eade92f7188d463089958699faf37c728b8a13e0` (and its validated ancestry) and
rerun the no-op scope, multi-family composition and metadata tests. Until that
integration rerun, this record does not claim final no-op scope qualification.
