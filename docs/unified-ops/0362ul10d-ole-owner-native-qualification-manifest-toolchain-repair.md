# 0362ul10d — OLE owner manifest and compiler-toolchain binding repair

- Date: 2026-10-24.
- Parent packet: 0362ul10c.
- Status: source-only false-acceptance repair; native execution and acceptance
  remain pending coordinator review.

The independent validator now invokes the paired runner's canonical manifest
policy loader before consuming any receipt.  It therefore rejects a changed
route set, ABBA/resource/allocator policy, or non-parameterized candidate
commit even when a caller recomputes the manifest hash in the receipts.  The
existing preparation checks still bind the selected candidate parameter to its
actual clean source worktree and commit.

Compiler selection is now source-bound and explicit.  The build phase selects
the absolute `cargo` and `rustc` dispatch paths from the effective `PATH`,
records each dispatch path, resolved regular-file path, byte count, SHA-256 and
version, and records the effective `PATH`.  Child builds and ABBA events set
`CARGO` and `RUSTC` to those exact paths.  Ambient Cargo/Rust compiler
overrides, `RUSTC_WRAPPER`, `RUSTC_WORKSPACE_WRAPPER`, `CARGO_BUILD_RUSTC`, and
`RUSTUP_TOOLCHAIN` are removed.  Every event records the complete toolchain
identity and sanitized environment; the validator compares those fields to
the build receipt.  Rustup dispatcher symlinks retain their basename for
correct dispatch while their resolved target is hashed.

Build-free adversarial coverage rejects both a validly rehashed alternate
rustc executable and a changed rustc version record.  Runner coverage uses
fake PATH tools with hostile ambient overrides and proves PATH selection and
wrapper removal.  Manifest mutations for a bogus route and `candidate_commit`
`deadbeef` are also rejected.

Source-only checks performed:

```text
python3 -m unittest tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run -v
python3 -m py_compile tools/native_qualification/*.py tools/test_ole_owner_paired_runner.py tools/test_validate_ole_owner_run.py
git diff --check
```

No Cargo build, native binary, timing result, or performance acceptance claim
is made by this record.
