# 0967 diagnostic evidence packet

This packet records the completed root-authorized 0967 capture for the exact
candidate scope below. The authoritative v2 read-only audit is complete:
the candidate decision is `reject_candidate`, `candidate_adoption` is false,
and there is no production recommendation.

The source revision is `893072657ad8b1d726bfce351b5ec2e98f084e94`. The candidate
scope is:

- `crates/litchi-doc/src/sprm.rs`
- `crates/litchi-doc/src/tracked_revision/codec.rs`
- `crates/litchi-doc/src/tracked_revision/package.rs`

The sealed terminal receipt is
[`receipts/capture-terminal-v1.json`](receipts/capture-terminal-v1.json), SHA-256
`d36312a1e6cef839a2ed3c67e7fc9b641fc264ed543a30f309300a6ebbfd55cc`. The runner exited 0
in root session `66783` and retained 896 packets: 128 control, 48 A/A, 144
allocator, and 576 native packet reports. A/A readiness passed its 1.05 bound
with 12 pairs per cell; the retained receipt is
[`analysis/aa-readiness-v1.json`](analysis/aa-readiness-v1.json).

The authoritative v2 raw audit is
[`review/independent-raw-audit-0967-v2.json`](review/independent-raw-audit-0967-v2.json),
SHA-256 `2712e39c456d7528c4c9347d6f085dd33253bf43d389d265b979ffee606b29ea`. It re-audited the same bound plan, manifest, retained
packets, binaries, and thresholds without a workload or measurement rerun. The
v2 decision is
[`review/independent-decision-0967-v2.json`](review/independent-decision-0967-v2.json),
SHA-256 `27ab3f1cdad3262a3d374d0dea83a065c2f15909ede165e96bc06518fc2cb1dd`. It records complete custody and valid oracles, but
rejects the fixed cohort after nine real native non-inferiority failures and
four RSS guard failures. Allocation benefit and the allocation-call guard passed;
those results do not qualify the candidate for adoption. The v2 helper is
[`review/independent_raw_audit_0967_v2.py`](review/independent_raw_audit_0967_v2.py),
SHA-256 `8215eaded6f6b0b6b45e42bf2fef603f5566c1ad0ec9c3a2b8fbfeca63451f77`.

The previously reported v1 sealed audit bytes were overwritten and could not be
recovered. The observed v1 artifact remains under
[`review/independent-raw-audit-0967-v1.json`](review/independent-raw-audit-0967-v1.json)
and the incident-specific preserved copy
[`review/incident/independent-raw-audit-0967-v1-overwritten-b756855c.json`](review/incident/independent-raw-audit-0967-v1-overwritten-b756855c.json),
both SHA-256 `b756855c7e779901be78d1c3574a5445e1a09078ebe25de0ffd410eb16fa3436`.
The v1 decision remains at
[`review/independent-decision-0967-v1.json`](review/independent-decision-0967-v1.json),
SHA-256 `f69d61f74e3943b9fee4411a1898a70aeb410e61702524b4f73c23bc54cb171f`,
and is historical with an invalid binding to the unavailable `dfb37103…` audit.
[`review/audit-custody-incident-0967-v1.json`](review/audit-custody-incident-0967-v1.json)
records the overwrite and its v2 resolution.

[`analysis/reduction-v1.json`](analysis/reduction-v1.json) and
[`analysis/selection-v1.json`](analysis/selection-v1.json) are the original
root outputs copied byte-for-byte. They retain their original
`pending_incomplete` and `no_decision` statuses, including the reducer's
incomplete classification of the timed allocator guard failure. The
authoritative v2 audit and decision classify the completed fixed cohort
without overwriting these original files; the v1 audit and decision remain
historical incident artifacts.

All raw capture files remain outside the repository under
`/home/zhuhe/litchi-performance-0967/capture`. The external inventory is
`/home/zhuhe/litchi-performance-0967/analysis/raw-capture-inventory-0967-v1.json`,
704038 bytes, SHA-256
`65dc6af75ae993a07890358bf12a1c97b65b365cce9092ad061f972bb9e2f6be`. It binds 3563 files totaling
610323073 bytes; the largest is 762426 bytes. The repository copy
[`provenance/raw-capture-summary-0967-v1.json`](provenance/raw-capture-summary-0967-v1.json)
records this inventory, the runner manifest, and the exact regeneration
reference. Raw stdout, stderr, RSS reports, packet JSON, build outputs, and
compiled binaries are not copied into this packet.

The diagnostic patch is retained under `repro/` with its isolated reproduction
receipt. The reproduction applies it with `/usr/bin/patch` in an explicit
scratch directory and does not touch the main checkout. The sealed control
harness source and lockfiles are also retained under `repro/`; the full source
roots remain external and are bound by the source manifests under
`provenance/`. `tools/run_0967_v1.py`, `tools/reduce_0967_v1.py`,
`tools/prepare_plan_0967_v1.py`, the plan binders, the scenario audit, and the
independent audit helper are copied for review and replay reference. No build,
workload, retry, threshold change, or production source write was performed
while assembling this packet.


Cleanup is recorded by [`receipts/build-cache-cleanup-v1.json`](receipts/build-cache-cleanup-v1.json) and its plan [`receipts/cleanup-plan-v1.json`](receipts/cleanup-plan-v1.json). The cleanup freed 10,343,804,928 disk bytes while preserving the 14 bound executables, source roots, and raw capture evidence; executable identities and hashes were unchanged.
`packet-manifest-0967-v1.json` lists every copied file, its size, and SHA-256.
