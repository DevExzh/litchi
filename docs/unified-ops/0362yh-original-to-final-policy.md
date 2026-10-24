# 0362yh — explicit original-to-final W8 projection policy

This record defines the reviewed protocol rebind for comparing the historical
f2 before endpoint with the candidate production source at
`c796d0acc24c7c6e6fd08980b9e31af6bb86b18b`.  The ordinary W8 runner remains
strict shared identity: `shared_identity()` still requires identical
projection, lock, probe and fixture identities.  The exception is opt-in via
`--original-to-final-policy` and `--original-to-final-receipt`; it does not
alter that generic path.

The checked-in policy is
`docs/unified-ops/results/0362yh/original-to-final-policy.json` (7,609 bytes,
SHA-256
`978d9c1b4eff2281295dca10b3b54b9fb93df1e07af7f342eb9cbfb865aae9b9`).  It
pins:

- original commit `f2bbbc8d07ab5da26c180f1ff74813d39f632d54` and final commit
  `c796d0acc24c7c6e6fd08980b9e31af6bb86b18b` (candidate production pin);
- the candidate source archive (`6,633,830,400` bytes,
  `65dc3102703cf9b4941003450440ac079bf74f269072f491cbce572e020bf9d8`),
  assembled without a production overlay;
- the exact target, release profile, toolchain, all-features selection,
  corrected probe (`75,035` bytes), fixture (`4,231` bytes), and 109-route
  order;
- the historical f2 lock (`139,441` bytes,
  `1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c`) and
  candidate lock (`139,606` bytes,
  `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`);
- the exact pre/post workspace and package manifests, toolchain, and complete
  dependency projections under `results/0362yh/images/`.

The graph proof recomputes both projections.  All 518 registry identities,
checksums and dependency lists must match.  Exactly six local package
dependency lists may differ: `litchi-iwa-core`, `litchi-odraw`, `litchi-ops`,
`litchi-ppt`, `litchi-xlsb` and `litchi-xlsx`; the `litchi-ops` delta includes
the reviewed `litchi-vba` feature edge.  The historical before all-features
selection predates the final `sign` feature and is pinned separately from the
candidate after selection.  The reviewed before source
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
The executable, build-free-to-native handoff is
`results/0362yh/c796-qualification-commands.sh`; it keeps targets, binaries,
receipts and ABBA output external, uses one offline locked build job, invokes
the complete untimed validator and pair receipt before guarded ABBA, and
fails closed for allocator because this probe has no allocator variant.

This is a protocol/provenance record only.  It performs no Cargo build, native
qualification, timing run, or inventory acceptance.  The `c796` source remains
a candidate production pin until the final source inventory review, endpoint
builds, untimed validation, pair receipt, and native ABBA review complete.

Review found that the analyzer still required the ordinary shared-identity
record after the runner had admitted the original-to-final mode. Follow-up
`c7a982b97f` adds mode-specific verification of the canonical pair receipt,
fixed endpoint artifacts, complete route order and recomputed projection
proof; ordinary analysis remains unchanged. `e814e4512f` registers the full
109-route, five-round regression in the existing projection-policy test
module. All three registered W8 modules now run 40 tests, including rejection
of source, projection and lock tampering. The reviewed packet merges at
`54398757b8`; these are synthetic analyzer tests, not native measurements.

Any receipt or binary made for 49c is stale and must not be relabeled as c796.
