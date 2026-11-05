# 0362ul31 — W8 File native build and 109-route validation

This is a reduced record of the actual W8 File build and correctness run.  It
is a provenance record, not a performance qualification or final acceptance
claim.  Large source archives, target trees, binaries and raw route files stay
outside Git under
`(deleted-coord-artifact)`.

The run compares the pristine historical f2 source plus the reviewed
consumer-only adapter with the immutable 472b production checkout:

| input | bytes | SHA-256 |
| --- | ---: | --- |
| f2 source archive | 6,614,528,000 | `de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280` |
| 472b source archive | 6,634,147,840 | `44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274` |
| f2 historical `Cargo.lock` | 139,441 | `1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c` |
| 472b `Cargo.lock` | 139,606 | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` |
| generated f2 consumer adapter probe | 74,719 | `3e14376c366fa911025212d7e7081ab5c3b5aaac804c1e01a1415fd5cd826ca0` |
| reviewed consumer-only adapter patch | 153,482 | `3e05d33946c7a9a2fea5c84cda33c799fbf8cd3c92fbcf890812736641ed01cd` |
| f2 dependency projection | 145,931 | `b070ddb516a0223f9f21509de6564c7c5690d26e955047a922d4d837b7522335` |
| 472b dependency projection | 146,621 | `2964946c802881909d5aaeea427fb32e85d0657f0987185ce21ff29e84c0d686` |
| shared corrected probe | 75,035 | `21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0` |
| shared 109-route fixture | 4,231 | `b7439d899a940b4135dd44081554a4aa9d7e70c19e2ccd3339440baad5ca7e09` |
| f2 route map | 264,946 | `28c9d78ab8c10cd4a18733a14c7c6c69c143efa45eebf26051b02878d8b4e4f8` |

The source commits are f2
`f2bbbc8d07ab5da26c180f1ff74813d39f632d54` and candidate production commit
`472b6705475b073e7fe08f6e11ee17fb1e29f838`.  The checked-in explicit
original-to-final policy is 7,609 bytes with SHA-256
`e56d549c733629ac73492a3511a45a1ed6a4cec97f1a4705ceea4ba9b67d97d2`.  It
requires all 518 registry identities/checksums/edges to match and permits only
the reviewed six local graph deltas.  The f2 and 472b workspace manifest hashes
are respectively
`66a8f4d2b0ffb6b9ba4eb87d84fee8a3049e8ef7e4380b6ee48acaaf372dd7b5` and
`f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`; their
package manifest hashes are respectively
`80b231a0a052f5848ca7358a870847672fe39d8a6fb56b03e8881b8dbd840db2` and
`004d3b3cba8c035d209665a76c9cfe3932f12a9c198a4aea9fa314ef8ebb3a39`.
The pinned toolchain hash is
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`.

## Commands and timeline

The external environment supplied `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, release debuginfo disabled,
offline mode and the memory-capped Cargo shim.  The shim is 520 bytes with
SHA-256 `1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`.
The recorded rustc artifact is 21,139,856 bytes with SHA-256
`5be191e8c70a5007074254cb74d78dbfdd25cd5a61c27de7b8e5410122ec24f4`.
Both builds used the x86_64-unknown-linux-gnu target, release profile,
`--offline --locked`, `-p litchi-ops --all-features`, and distinct fresh
external target directories.

The build and initial validation used the pre-repair qualification recipe
(the recipe at `87377c1801`, whose endpoint arguments still named the
validation directories):

```text
source (deleted-coord-artifact)
export W8_PREP=(deleted-coord-artifact)
export W8_BASELINE_ROOT=(deleted-coord-artifact)
export W8_BASELINE_ARCHIVE="$W8_PREP/baseline/f2-source.tar"
export CANDIDATE_SOURCE=(deleted-coord-artifact)
export CARGO=(deleted-coord-artifact)
bash docs/unified-ops/results/0362yh/472b-qualification-commands.sh build
bash docs/unified-ops/results/0362yh/472b-qualification-commands.sh validate
```

The build session was `43584`.  Both release builds completed and both
validators completed all 109 routes.  The initial `validate` then stopped at
endpoint assembly because `validate.py` writes a directory containing
`receipt.json` and `routes.jsonl`, while the pre-repair recipe passed the
directory itself as `--validation-receipt`.  No endpoint manifest was claimed
from that failed step.  The two complete validation receipts and route reports
were retained; rerunning `validate` after the repair correctly refused to
overwrite them.

The endpoint recipe repair is commit `2a3ace8002` (merged as
`553dab051748dd67247cf3909bfa4dc11fa357f3`).  It binds
`validation-before.json/receipt.json` and `validation-after.json/receipt.json`.
Using that repaired recipe, the two endpoint commands were regenerated
directly from the retained validation outputs, without rebuilding or rerunning
the routes.  The subsequent `pair` command used the repaired recipe and ran
the complete endpoint/policy validator.  The pair receipt status is
`verified_protocol_only`.

## Run artifacts

The build receipts bind the source archive, source assembly, binary, lock,
probe, fixture, projection, manifests, toolchain, Cargo shim, compiler,
target, profile and exact command:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `build-before.json` | 2,881 | `74810d2fcc4b6b4ac5f1341bf3fa7f2b7892c90fec8261fb0be0999148fc6b7a` |
| `build-after.json` | 2,834 | `40acbe826aa3e12069e53e36e4dba4d00db75c11f336d8152ff2b6f22826c511` |
| f2 binary | 67,879,576 | `be7e3dbd9c31403d3bc017c07299f391e2b4b5b8f889107f9e0bfd686ada75c6` |
| 472b binary | 70,239,360 | `7c9061725c74c2fc255d33694cbbe1f6de0dd5f750cc8d1017dc82811bef58f6` |

The path-bound source assemblies used by those receipts are the external
744-byte f2 assembly (`338c8c669dbb1114566bee1197a5946489caba0fee62c5952e7f35d55cd23f11`)
and 483-byte 472b assembly
(`cc9d9f67e8d4474685fa0daf4ba9054adcb508ca6dcc047f49fda39aa6b1289f`).
The f2 assembly records the reviewed adapter patch and unchanged 459,676-file
production tree; the 472b assembly records no production overlay.

Both validators report route count 109.  The route reports are identical
238,193-byte JSONL artifacts with SHA-256
`877802466f1a0b4338b57e58a2cc9ad614427869ad4eef69e1bf7f6b78f71761`.  Their
canonical validation receipts are:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| f2 validation receipt | 38,146 | `e25498dfe36da33949a4e6bf8ec8b44588864ebf27f43e50f781ab7e081f0dfc` |
| 472b validation receipt | 38,150 | `35b72f6ee5aef761b440b824d8119a9533f5df35b7821618d24907dd3e29aaeb` |

The endpoint manifests are:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `endpoint-before.json` | 3,193 | `72080e5313b5ba183a45769d2e78a9116d8b3ea7527f7af8b9aae059f42e2633` |
| `endpoint-after.json` | 3,152 | `500c156c1fe1b9962a94ff0f5f8c03752ca1fdd08b3725438f2811c52466ab63` |

The original-to-final pair receipt is 8,949 bytes with SHA-256
`f59c22c65737495b7adc88a5c0c6286680f92ed4e101009f66ca81e4c3d72275`.
It binds the two source commits and archives, both locks and projections,
the manifest identities, all-feature release selections, target/toolchain,
109-route order, and the six permitted local graph deltas.  It is a protocol
receipt only; it does not claim native timing, ABBA, allocator or final source
inventory acceptance.

The no-Cargo focused suite was run after the repair:

```text
python3 -m unittest tools.test_w8_projection_policy \
  tools.test_w8_file_provenance tools.test_w8_original_baseline
```

It passed all 45 tests.  `bash -n
docs/unified-ops/results/0362yh/472b-qualification-commands.sh` also passes.
No binaries or raw evidence from this run are committed.
