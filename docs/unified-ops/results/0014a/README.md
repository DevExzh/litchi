# Record 0014a probe evidence

Two temporary probes, built, run and deleted during record 0014a (as record
0013 §2 did). Both reuse the `families()` table and helpers of
`crates/litchi-pptx/tests/package_signature_policy.rs` at the record's commit:
the probe file was that test file with its tests replaced by the probe below,
saved as `crates/litchi-pptx/tests/zz_*.rs`, and run with

```sh
CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 \
  CARGO_PROFILE_TEST_DEBUG=0 cargo test -p litchi-pptx --all-features \
  --test zz_<probe> -- --nocapture 2>&1 | grep '^REPORT\|^UNSIGNED'
```

"Base" means the litchi-pptx sources of `0b79b2f61a` (`git stash push --
crates/litchi-pptx/src`, with the new `package/signature.rs` moved aside);
"after" means the record's commit.

## `base-signed-families.txt` (base sources)

For each producer-signed fixture and each family of the table: the family's
exact no-op and its changed edit on `Package::from_vec(signed bytes)`, then
`Package::opc()?.is_signed()` and `Package::to_bytes()`:

```rust
let result = (change.edit)(&mut package);
let signed = package.opc().map(|opc| opc.is_signed()).unwrap_or(false);
// "source" when to_bytes() reproduces the signed input,
// "changed(signature=..)" when it publishes other bytes.
```

Result: 40 of the 44 changed edits (20 families on 2 fixtures) succeeded,
removed the signature (`signed_after=false`) and published a package without
signature parts: the silent strip. The 4 others (a raw `Package::edit_opc`
change and the opened-presentation transaction, on 2 fixtures) were already
refused. Every no-op of the table was already exact; the no-ops that were not
(class c) are pinned by the dedicated tests of the test file, which fail on
the base sources (`legacy_comment_no_ops_keep_a_signed_package_exact`,
`modern_comment_no_ops_keep_a_signed_package_exact`,
`modern_comment_command_no_op_keeps_a_signed_package_exact`,
`theme_no_ops_keep_a_signed_package_exact`).

## `unsigned-outputs.txt` (after; identical on base)

The same no-ops and changed edits on unsigned packages: each fixture after
`Package::unsign()` and `to_bytes()` (prepared like the test where the family
has a `prepare` step), then the class c legacy-comment no-ops and one changed
reorder on an unsigned package holding comments. Each line gives the result
and the output length and 64-bit FNV-1a hash:

```rust
fn fnv(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
    hash
}
```

The base and after runs printed byte-identical reports (145 lines, SHA-256
`e0a092a8bbff5624f327f57e53118ed9699fd08cdacf71c4cdd103ae98540938`), so the
94 published outputs of unsigned packages, no-ops and changes, are the same
bytes before and after the change.
