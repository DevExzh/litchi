# 0968 bounded diagnostic evidence packet

**Status: sealed bounded diagnostic packet; v6 capture, independent raw audit, and cleanup pass.**
This packet is diagnostic evidence only. It contains no raw capture children,
profiler streams, source export, target directory, or compiled executable.
Every copied file is below 1 MiB.

The retained CPU record is the offline attribution of retained 0942
`cycles:u` callchains in
[`analysis/v2/cfb-commit-attribution-v1.json`](analysis/v2/cfb-commit-attribution-v1.json)
and its Markdown companion. It makes no 0968 build, workload, latency,
allocation, speedup, causality, or adoption claim. The source provenance is
bound to revision `6493e26155ee05e554f1371dd2fad707cd4bca0e`; the current source
seal remains [`protocol/seal-v4.json`](protocol/seal-v4.json). The source patch,
build configuration, source review, output oracle, and CPU review are retained
in the packet with their hashes in the manifest. The original seal validation
receipt [`review/seal-v4-validation-v1.json`](review/seal-v4-validation-v1.json)
is preserved unchanged. Its historical superseded-seal-v3 preservation field
was 63 characters; [`review/seal-v4-validation-correction-v1.json`](review/seal-v4-validation-correction-v1.json),
SHA-256 `44fa159ec9c1fdd6f4776be6961c5979f47038a62b8bddf6304ce75300cbfbee`,
records the full 64-character seal-v3 hash. This correction affects only that
historical descriptor; current seal-v4, source, executable, plan, capture, and
audit bindings are unchanged.

The corrected census protocol and v6 evidence are linked as follows:

- [`protocol/plan-0968-prospective-v6.json`](protocol/plan-0968-prospective-v6.json)
  is the prospective no-build/no-workload plan. Its SHA-256 is
  `3ffcc37e7078f716082844ed0689bdfaf500e820ec5d2ca11a1c61bd47638797`.
- [`protocol/protocol-seal-0968-v3.json`](protocol/protocol-seal-0968-v3.json)
  binds the corrected per-edit cardinality rule and the 48-child matrix.
  Its SHA-256 is
  `bc970a5f03fb45ec8c6257d1ea452b71db2d4e1babfa3a79dcd3135449a81a42`.
- [`receipts/root-go-v2.json`](receipts/root-go-v2.json) is the explicit
  no-retry diagnostic GO. Its SHA-256 is
  `3ec3b54400bea1bef26522a974fb7494d66d58f2193a71913aff02e2371c5ab3`.
- [`analysis/census-summary-v1.json`](analysis/census-summary-v1.json) is the
  v6 reduction: 24 cells, 48 packet reports, two repeats per cell, zero
  capture/commit/layout errors, and decision `diagnostic_only_no_adoption`.
  The reduction SHA-256 is
  `48fc947ac31ce136ea45f3c2b55143e00543e65ad6accd02f054424ab4304035`.
- [`review/independent-raw-census-audit-v2-final-v1.json`](review/independent-raw-census-audit-v2-final-v1.json)
  is the final independent raw audit. It passes custody, schema, stage
  conservation, fixed schedule, ordinary-output SHA/length, all 20
  correctness gates, and equality of all 24 repeated cells. It makes no
  timing, allocator-total, speedup, causality, or adoption claim. Its SHA-256
  is `1fb1c5b5aaad32d8b55770d3e06917927b7cc6e4f7b7100ae02b2d58f6bb0333`.
  The bounded auditor is retained at
  [`tools/independent_raw_census_audit_0968_v2.py`](tools/independent_raw_census_audit_0968_v2.py).

The root GO binds the restored original protocol approval
[`review/protocol-approval-v2.json`](review/protocol-approval-v2.json) at its
exact SHA-256
`2d8468789f99de61042d74183ce8e680e1326d4d213e43f13995ae7d54150225`.
The amended approval is preserved separately as supplemental history in
[`review/protocol-approval-v3-amended-retained-v1.json`](review/protocol-approval-v3-amended-retained-v1.json),
SHA-256
`dbd51169792a19350b2c308b6eb0d21746a72be08645bfab3ac9161a618bc70c`.
The custody resolution is
[`review/protocol-approval-custody-incident-v1.json`](review/protocol-approval-custody-incident-v1.json),
SHA-256
`ed51390508d8795432637c954d2eeb58a9d9536dc8ebd865be646725056cd949`.
The original v2 approval remains the binding artifact; the amended v3 bytes
are not substituted for it.

The first v5 epoch is retained only as incident history. The checker rejected
36 packets because it assumed one commit/capture attempt per child even though
multi-edit rows emit one successful stage record per requested edit. All 48
processes terminated with return code 0, but v5 is excluded from formal
evidence. The bounded incident records are
[`review/census-v2-incident-v1.json`](review/census-v2-incident-v1.json) and
[`review/incident-census-v5-per-edit-count-v2.json`](review/incident-census-v5-per-edit-count-v2.json).
The corrected v3 helpers and their unit test are under `tools/`; their
source paths, sizes, and SHA-256 values are recorded in the protocol seal and
packet manifest.

The capture manifest remains external at
`/home/zhuhe/litchi-performance-0968/capture/census-v6/capture-manifest-v3.json`
(94,568 bytes, SHA-256
`7f36d3ed05d89ab1ede777703c27536518116c47e8bef91955396c573ed715d9`). Raw
capture children and stdout/stderr streams remain external and are not copied
into this repository. The external diagnostic executable is 55,125,224 bytes
with SHA-256
`50693dd8d2e7562f69c86d08872246f1a5d040be9f4f42fc850ed9c0a309dac9`.
[`provenance/census-v6-binding-0968-v1.json`](provenance/census-v6-binding-0968-v1.json)
binds these external descriptors without importing them.

The cleanup plan is retained at
[`receipts/cleanup-plan-v1.json`](receipts/cleanup-plan-v1.json). The completed
cleanup receipt is [`receipts/build-cache-cleanup-v1.json`](receipts/build-cache-cleanup-v1.json),
SHA-256
`0c45dde744f4f0488c0e737a6bdd46e1321db8f800b8d506ce63a0b956b8d5ee`. It
records 8,896,917,504 allocated disk bytes freed while preserving the bound
release executable byte-for-byte and preserving source, fixtures, raw capture,
and review evidence. No production source, build, workload, or capture action
is authorized by this packet.

For bounded helper verification, run
`PYTHONDONTWRITEBYTECODE=1 python3 tools/test_cfb_commit_census_0968_v3.py`.
Replaying the reduction or audit requires the external v6 capture root and a
fresh output path; do not overwrite the canonical summary or audit files. The
full source, binary, raw capture, and regeneration commands remain at
`/home/zhuhe/litchi-performance-0968`.

The coordinating record is
[`../../0968-cfb-commit-census.md`](../../0968-cfb-commit-census.md). The
sealed packet inventory is
[`packet-manifest-0968-v1.json`](packet-manifest-0968-v1.json).
