# 0362ul79a — receipt-bound raw-census validation

## Scope

The 0362ul79 plan described the 84-profile final census, but a rustdoc path
alone did not prove which package configuration produced that input. This
follow-up adds the receipt binding required before final `--require-complete`
can be accepted. It is source-only: no Cargo, rustdoc, native build, or
performance measurement was run, and no generated rustdoc or receipt artifact
is committed.

## Receipt contract

`tools/raw_api_coverage.py` accepts one `--receipt PROFILE=PATH` binding for
each declared profile. A receipt has an exact schema (`schema=1`) and records
the profile/package, default-feature boolean, ordered feature list, exact
`cargo rustdoc` argv, source commit and tree, source-root path, external
target directory, the approved root `Cargo.lock` path and SHA-256, the selected
root toolchain path and SHA-256, Cargo and rustc executable paths and SHA-256,
and the rustdoc artifact path, byte count, and SHA-256. The manifest also
records the coordinator-approved seeded-lock SHA-256; a receipt cannot choose
a different ignored lockfile.

Validation binds every field to the manifest and supplied rustdoc input:

- profile, package, default features, feature order, and argv must match the
  declared profile exactly;
- the receipt source commit must equal both the manifest revision and current
  `HEAD`, and its recorded tree must resolve from that commit;
- source root must be the audited checkout; target, artifact, receipt, Cargo,
  and rustc paths must be absolute external paths (not under the checkout or
  `/tmp`), with the target directory present;
- the selected root toolchain and exact coordinator-approved seeded `Cargo.lock`
  must match their recorded hashes; and
- the supplied rustdoc file must match its recorded bytes and SHA-256.

This rejects stale captures and captures made with a different feature graph
even when the resulting crate surface happens to have the same shape.

## Source-overlay policy

Before validating a receipt, the checker inspects tracked, untracked, and
ignored checkout state. Every overlay fails closed, including ignored source,
manifest, fixture, target, and vendor trees; ignored directories are walked
recursively so a complete vendor tree cannot be hidden behind one directory
status line. The sole in-checkout exception is the exact root `Cargo.lock`
whose SHA-256 matches both the receipt and the coordinator-approved manifest
hash. Rustdoc, receipts, target data, and tool binaries remain external
scratch paths.

The coverage manifest now records trusted Git blob identities for its expected
profile-source and workspace-source paths. Those paths are verified through
Git object type and blob SHA, so a sparse checkout need not materialize them.
Callable anchors remain source- and declaration-bound, and final mode still
requires human/source-reviewed dispositions; no semantic or owner exemption is
inferred automatically.

## Source-only checks

```text
python3 -m py_compile tools/raw_api_coverage.py tools/test_raw_api_coverage.py
exit 0
python3 -m unittest tools.test_raw_api_coverage
Ran 14 tests in 0.245s ... OK
python3 tools/raw_api_coverage.py --root . \
  --manifest tools/raw_api_coverage.json \
  --glossary tools/api_glossary.json --json
exit 0; complete=false; source_failures=[]; current_source_failures=[];
missing_receipts=15; missing_inputs=15; missing glossary profiles=53;
missing_workspace_packages=0
```

The normal report remains deliberately incomplete because final rustdoc and
receipt inputs have not been captured. The three receipt regressions exercise
valid identity binding, feature and artifact tampering, stale source commits,
exact duplicate receipt rejection, and ignored vendor/manifest/fixture
overlays plus wrong seeded-lock hashes. The fixture creates its source tree,
tool shims, target, artifact, and receipt outside `/tmp`; it is deleted by the
test teardown.
