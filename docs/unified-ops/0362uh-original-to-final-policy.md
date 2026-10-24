# 0362uh — explicit original-to-final W8 projection policy

This record defines the reviewed exception for comparing the historical f2
before endpoint with the f6fe final endpoint.  The ordinary W8 runner remains
strict shared identity: `shared_identity()` still requires identical
projection, lock, probe and fixture identities.  The exception is opt-in via
`--original-to-final-policy` and `--original-to-final-receipt`; it does not
alter that generic path.

The checked-in policy is
`docs/unified-ops/results/0362yh/original-to-final-policy.json` (7,307 bytes,
SHA-256
`5f198d0e4462529783538b6a7fa2f5a8cab025bc191f9ce582f150a9cd696d2e`).  It
pins:

- original commit `f2bbbc8d07ab5da26c180f1ff74813d39f632d54` and final commit
  `f6fe430a4d19d94aa360c0e9ed03ebd757265225`;
- the exact target, release profile, toolchain, all-features selection,
  corrected probe (`75,035` bytes), fixture (`4,231` bytes), and 109-route
  order;
- the historical f2 lock (`139,441` bytes,
  `1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c`) and
  f6fe lock (`139,591` bytes,
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

`tools/w8_projection_policy.py pair` creates a canonical pair receipt after
checking both complete endpoint manifests.  The ABBA runner consumes that
receipt after its normal endpoint validation and rechecks the small fixed
projections and source-assembly metadata without hashing either multi-gigabyte
source archive a second time.  Route order and all endpoint artifact
identities remain bound to the receipt.

This is a protocol/provenance record only.  It performs no Cargo build, native
qualification, or timing run.  The f6fe pins are diagnostic until the final
production source freeze is selected; a later final source commit requires a
new reviewed policy and matching pre/post images.
