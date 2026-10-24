# 0362yh — explicit original-to-final W8 projection policy

This record defines the reviewed protocol rebind for comparing the historical
f2 before endpoint with the candidate production source at
`49c25ef67f62a8bb943e7c6a0bf464e49ae87d07`.  The ordinary W8 runner remains
strict shared identity: `shared_identity()` still requires identical
projection, lock, probe and fixture identities.  The exception is opt-in via
`--original-to-final-policy` and `--original-to-final-receipt`; it does not
alter that generic path.

The checked-in policy is
`docs/unified-ops/results/0362yh/original-to-final-policy.json` (7,307 bytes,
SHA-256
`afb0d053a98ddb3d26d50d1d90944aab0fcc6c536c2cd99024162d49642d3d7f`).  It
pins:

- original commit `f2bbbc8d07ab5da26c180f1ff74813d39f632d54` and final commit
  `49c25ef67f62a8bb943e7c6a0bf464e49ae87d07` (candidate production pin);
- the candidate source archive (`6,630,123,520` bytes,
  `127d81f130c6782c9329844b0992adb5b76b5f5e7d5258f3d95d1360e620db10`),
  assembled without a production overlay;
- the exact target, release profile, toolchain, all-features selection,
  corrected probe (`75,035` bytes), fixture (`4,231` bytes), and 109-route
  order;
- the historical f2 lock (`139,441` bytes,
  `1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c`) and
  candidate lock (`139,591` bytes,
  `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`);
- the exact pre/post workspace and package manifests, toolchain, and complete
  dependency projections under `results/0362yh/images/`.

The graph proof recomputes both projections.  All 518 registry identities,
checksums and dependency lists must match.  Exactly six local package
dependency lists may differ: `litchi-iwa-core`, `litchi-odraw`, `litchi-ops`,
`litchi-ppt`, `litchi-xlsb` and `litchi-xlsx`.  The reviewed before source
assembly is the f2 archive plus the consumer-only compatibility adapter; the
after assembly has no overlay.  No caller-supplied hashes, allowlists, source
overrides or candidate-manifest overlay can replace these pins.

`tools/w8_projection_policy.py pair` invokes the existing complete endpoint
validator for both manifests before creating a canonical pair receipt.  The
receipt binds both source commits and archive identities, every manifest/lock/
projection/probe/fixture identity, the complete build selection, and the exact
route order.  The ABBA runner uses the same complete validator in pair mode,
then rechecks the fixed projections and source-assembly metadata, including a
second hash of each multi-gigabyte source archive, before timing begins.
Adapter patches, source-assembly records, binaries, build receipts and untimed
validation receipts therefore cannot be replaced by receipt-only metadata.
That startup cost is outside every measured route and prevents a missing or
mutated source from passing on metadata alone.

The reduced build-free rebind receipt is
`results/0362yh/final-source-preparation.json`.  It records the regenerated
candidate fixture/projection, the exact six local graph additions, and the
stale f2 endpoint and b673 candidate-lock inputs that are explicitly rejected.

This is a protocol/provenance record only.  It performs no Cargo build, native
qualification, timing run, or inventory acceptance.  The `49c` source remains
a candidate production pin until the final source inventory review, endpoint
builds, untimed validation, pair receipt, and native ABBA review complete.
