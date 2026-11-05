# 0361l — Canonical iWork detection and legacy bundle ingress

- Date: 2026-10-22.
- Base: `50122ae80aa68d9d80cde4d722db7b37fdcebb4f`.
- Contract: 0351 §2.5; follow-up to 0361d and 0362b.
- Status: implementation complete on `uo2/iwa-detection-api`; coordinator
  integration and the final workspace gate remain pending.

## API migration

The public raw iWork detector has one entry point per operation, with a required
`litchi_iwa_detect::Limits` argument:

* `PreparedSource::{from_bytes,from_shared_bytes,from_frozen_entries,from_path}`
  now take `Limits`.
* `litchi_iwa_detect::{bytes,reader,path}` now take `Limits`.
* The former public `_with_limits` twins and their default wrappers are removed.

Former default-policy consumers pass `Limits::default()` at their call site.
Existing explicit profiles are passed through unchanged; the Keynote presentation
route keeps its `source_limits.detector_limits()` profile. The facade and smart
detector explicitly select the detector's bounded default because they had no
caller-selected detector profile before this migration.

`litchi_iwa::Bundle::{open,from_bytes}` now require `BundleLimits`. The legacy
`Document` owner keeps its public default ingress, while its explicit helper stays
crate-private and passes the required bundle profile to `Bundle`.

`PackageLimits::new_with_limits` is renamed to `PackageLimits::from_values`; all
five dimensions remain configurable and the three-dimension `PackageLimits::new`
convenience constructor retains its previous fixed input/stream ceilings. No
ArchiveInfo or retained-constructor APIs are changed; those belong to the
iWork-finish slice.

The two generated Keynote chart-title `new_with_limits` aliases are removed.
`with_output_and_title_limits` remains the sole constructor and no workspace
caller used either alias.

## Consumers

The migration covers the detector crate's tests, Pages/Keynote/Numbers prepared
source tests, the iWork ops adapter, the facade's iWork tail and detection tests,
the legacy bundle/document tests, object-index and chart test helpers, and the
Keynote explicit source-policy route. The glossary gains an `iwa-detect` profile
with the `_with_limits` suffix ban; the existing `iwa` profile continues to cover
the host crate.

No parser, classification, source ownership, cursor restoration, archive proof,
allocation, or error behavior changes. In particular, detection still rejects
over-limit input before archive admission, restores reader cursors, preserves
the retained prepared-source profile, and leaves owner-specific source limits
authoritative when a leaf reopens under a non-default policy.

## Validation

Validation commands and results are recorded in the coordinator handoff. The
affected IWA, detector, proto, leaf-consumer and facade targets must pass strict
formatting, warnings-denied Clippy, consumer compilation and the glossary/API
checks before integration. Build artifacts are kept outside the checkout under
the capped task target directory; no benchmark binary or large performance
evidence is committed.
