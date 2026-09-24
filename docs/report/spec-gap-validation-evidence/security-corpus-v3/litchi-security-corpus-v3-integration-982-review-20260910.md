# Security corpus V3 current-982 integration review — 2026-09-10

## Scope and result

Read-only review of `/var/tmp/litchi-security-corpus-v3-integration-982-20260910`, compared with the frozen `ff954` overlay and the current-982 provenance/preimage records. The feature/source overlay is coherent; one publication-evidence hash is stale and must be refreshed before calling the evidence receipt final.

Candidate HEAD is `9821676451bf098cd0794fd177702bd82f27d826`. Its status is the expected eight modified feature source files plus the security corpus and evidence additions. `git diff --check` passes. No candidate, primary, index, stage, or commit was changed by this review.

## Source closure and merge checks

- Rehashed all 49 entries in `docs/performance/results/security-corpus-v3/integration-source-manifest.json`: 49/49 size and SHA-256 checks match.
- Required closure paths are present, including `litchi-crypto/src/spaces.rs`, OPC limits/physical package, and DOC package model/document package.
- The independent three-way editor merge matches exactly:
  - candidate `crates/litchi-ole-common/src/object/editor.rs`: SHA-256 `0e3d136ff167f82cb960f37766e996b2aa1f6450008d2b4ba364814e69a42542`
  - `/var/tmp/litchi-root-xls-security-editor-merge-20260910/merged.rs`: the same SHA-256
- XLS bounded admission remains in the eight-file delta. `Snapshot::from_bytes` delegates to bounded defaults; explicit object and CFB limits route through `PackageEditor::open_with_cfb_limits`; the current CFB visitor APIs remain present and bounded in the current-982 closure.
- The current overlay `crates/litchi-xls/src/cell_values/mod.rs` is byte-identical to the prior reviewed security payload, so the XLS admission/visitor integration was not silently replaced by current pending work.

## Lockfile

Parsed `tools/perf-baseline/Cargo.lock` against current candidate HEAD:

- 238 package records before and after
- zero added or removed package records
- zero version, source, or checksum changes
- only four dependency-array changes: `aes -> zeroize`, `litchi-crypto` feature dependencies, `litchi-perf-baseline -> litchi-crypto`, and the authoritative `litchi-xlsb -> litchi-xldm` edge
- candidate lock SHA-256: `577ca67feffe8866fb061b82dca0e36203d66f72e913fb7861c2332e38427562`

No stale or broad registry upgrade is present.

## Evidence integrity finding

The external publication receipt `/var/tmp/litchi-security-corpus-v3-integration-982-publication.json` lists 60 paths. Rechecking every listed path against the frozen overlay found 59 matches and one mismatch:

- receipt claims `docs/performance/results/security-corpus-v3/integration-provenance.json` is 6,968 bytes with SHA-256 `813d3b6057c273f78faaf39c96e7c89b56221d6cf0916bcb947e57f72701e28c`
- frozen overlay currently contains 7,676 bytes with SHA-256 `f08ad864c3259bf75ce7c613fd9fc2e75361a0c8f03b5f9f3e84ce8081cb5b7d`

The provenance file was modified after the receipt timestamp and explicitly says its final hash is recorded by that receipt. The feature payload hashes, source closure, preimages, and editor merge remain consistent; regenerate the publication receipt after the provenance file is frozen before publication.

The current provenance’s own SHA-256 is `f08ad864c3259bf75ce7c613fd9fc2e75361a0c8f03b5f9f3e84ce8081cb5b7d`; its source manifest SHA is `a9216e527e14e4b38d22f35690959c139f36ee2348eece646d9b24e4c91e7ab9`, and preimages SHA is `8ce5b6b246d14e3cbff93194ab20282a6338bff7218365b59eba8b7e7cfa205b`.

Root’s 7 binary tests, 66 correctness rows, and strict Clippy run were not rerun here per the handoff; this review checked the frozen source/provenance and the lock/merge invariants only.
