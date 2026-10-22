# 0362o — DOCX/PPTX metadata commit triad

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

## Coordinator integration

The coordinator assembled the six shared-properties repairs listed in the
[receipt](results/0362o/validation.json) over `51c7fbb1b5`. The repaired scope
code matches `9f9fa3930a` exactly after preserving the 91 lines of W4b deferred
snapshot helpers already present in the assembly. The XLSX adapter and
transaction tests match that worker revision exactly.

This closes the correctness return in 0360ab: no-op scopes authenticate the
recorded property state, compatible complete/partial scopes compose, and
creation/removal inverses work with independently captured no-ops. Supplemental
absence guards cannot hide the actual creation relationship ID or the
removal's inbound-relationship preconditions. All 17 shared transaction tests
pass, including those regression cases and eager/source-backed refusal checks.

Fresh retained integration validation passes:

- common OOXML library: 472 tests;
- XLSX metadata filter: 51 tests;
- DOCX metadata/core-properties/signature policy: 42 + 2 + 9 tests;
- PPTX metadata/families/signature policy: 120 + 6 + 10 tests;
- all-target/all-feature Clippy for common OOXML, XLSX, DOCX, PPTX and ops;
- workspace formatting.

The receipt records every command, exit status, elapsed time and external log
size/SHA-256, plus source hashes and the retained W4b helper proof. Commands use
the coordinator's 20 GiB capped environment, one Cargo job and four test threads.
The earlier worker-terminal results above remain separate from this fresh run.

The record is renamed from 0362n to 0362o to avoid the ODF raw audit's reserved
number. This is a committed integration candidate, not performance acceptance
or a completed final workspace gate. Native performance qualification and
remaining integrations are still required before trunk acceptance.
