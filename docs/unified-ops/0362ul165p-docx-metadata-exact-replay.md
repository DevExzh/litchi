# 0362ul165p — Exact DOCX metadata replay

DOCX metadata patches now retain the exact core-properties delta. Applying
an inverse restores the original core XML instead of regenerating equivalent
typed values. Application/custom properties are planned once against the
source and published first; core properties are then planned against that
intermediate graph. Replay follows that order, and inverse replay reverses
it. The public `Patch::parts()` remains the source-scoped application/custom
patch actually used for publication.

Changed replay preserves `metadata::Error::Signed`, and application/custom
failures preserve `Error::Properties`. Exact no-ops avoid an unnecessary OPC
clone. Review returns for intermediate-graph public patches, collapsed error
classification and duplicate application/custom planning are repaired rather
than accepted as exceptions.

The shared OOXML planner also keeps every base relationship ID reserved for
the entire plan. Previously, removing application properties could free an
ID for custom-properties creation in the same patch, although patch
validation authenticates every change against the original graph. Such a
patch could not replay. A regression now proves forward publication and
byte-exact inversion of that removal/creation pair.

## Fidelity and validation

The original opaque-core and inverse whole-archive assertions remain strict.
A new forward test compares replay from a reopened source with publication
from an authored package. The OPC writer deliberately preserves source-member
order for the former and uses canonical order for the latter; its existing
`topology_part_add_preserves_source_and_appends_deterministic_members` test
specifies that behavior. Independent review confirms that whole-ZIP equality
at this cross-provenance seam was too strong.

That forward test now compares unique member names, raw local records
including compressed payloads and descriptors, decoded payloads, central
records with only local-header offsets normalized, archive comments, ZIP64
mode and semantic projections. Count and raw-name alignment assertions
prevent a truncated or mismatched comparison. A negative control detects a
changed payload. The inverse test still requires the original archive bytes.

Root merges leaf `e7f9e97f0a6312adf7ed28910fd01a4f594dd724` after:

- **532 shared OOXML library tests pass**, with all features. Those source
  files are unchanged between the tested `f739e2e28f` and the merged leaf.
- **1,532 DOCX library tests pass**, with all features, at the merged leaf.
- Root verifies source/lock stability, log and Cargo-target trace hashes,
  terminal process state and target usage below the 20 GB limit.

The commands are:

```sh
cargo test --offline --locked -p litchi-ooxml-common --all-features --lib --no-fail-fast
cargo test --offline --locked -p litchi-docx --all-features --lib --no-fail-fast
```

These are affected-library results, not final all-65 or native acceptance.
The existing 28 native groups do not measure the changed DOCX metadata
commit/apply paths. A focused 29th group is required for core, application,
custom, combined and no-op commits/replay, with separate allocation and ABBA
measurements. Its preparation is in progress; performance is not waived.

## ODF consumer and cleanup status

The original `70d` ODF probe reaches Rust after complete tracked source
materialization, but the broad benchmark package pulls in eleven unrelated
OOXML facade compiler errors. An isolated consumer is being reviewed. It
retains all ten ODF families and facade routes, the canonical probe and the
optimized profile, while leaving other benchmark consumers unchanged.
Endpoint-specific manifests preserve the historical registry compression
dependencies and the candidate's vendored dependencies. Independent baseline
lock provenance is required; a supplied candidate lock is not an acceptable
baseline. Compiler and native qualification remain pending.

The merged quarantine-repair leaf has been retired after clean/holder checks,
bringing cumulative cleanup to 280 worktrees and 442 refs. The census leaves
all 70 excluded refs unchanged. Required validation inputs remain retained;
completed repair leaves are retired after their merge. Nothing is pushed.

## Evidence

Paths are relative to `(deleted-coord-artifact)`.

| Artifact | SHA-256 |
| --- | --- |
| `docx-metadata-7e2b-independent-review-20261026/review-receipt.json` | `041cea413e6308f15fb213667d9ee0d548f99aa6acec2be00b8005a2fa130580` |
| `docx-metadata-full-06394-order-diagnosis-20261026/review-receipt.json` | `3759a3bfe6ced72f93bdc0e84ff57a3eb500d19fb181d2088203d891ee3f6e4c` |
| `docx-metadata-06394-order-root-diagnosis-20261026/receipt.json` | `fbb3927df4968d3b82fc021422e2f1d5f8cf7cdfe222d4d2a217ef31b40f99c8` |
| `docx-shared-f739-root-verification-20261026/receipt.json` | `2ca5dde2d947002930b16cacee9d7c48d4cf91680df38c9515b8b37071b9e184` |
| `docx-metadata-e7f9-root-verification-20261026/receipt.json` | `701fbaee3a83e555cb6f86932acdf064299ec581c864480631e84ec0c0972fe4` |
| `odf-raw-original70d-v5-root-verification-20261026/receipt.json` | `30b4a1897d89d14f76c1b1986986a4527f371de84bdefa8b7cff28c90ce6d658` |
| `gate-quarantine-scope-retirement-20261026/receipt.json` | `dda324d48641a25ac81c5900599e8891e80f2a0b7cb9488cb399710754ffa83e` |

Failed diagnostics and their logs remain preserved. Memo 22 alone is deferred;
the final gate, candidate/protocol rebinding and all native qualification
remain required.
