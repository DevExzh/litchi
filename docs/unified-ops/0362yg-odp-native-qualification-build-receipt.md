# 0362yg — ODP native qualification build boundary

This is a reduced execution receipt for the reviewed 0362yf harness. It is
not a native timing qualification and does not replace the accepted 0362yd
baseline evidence.

## Frozen inputs

- Candidate source provenance: `f28a599077c8537c2de729a5c5cb2a4e1f85ec3d`,
  with the probe-only fixes now recorded in `1b3ec77fa5`.
- Historical baseline provenance: e4a
  `e4a0614f5d3a1d768346cc33a4a3b86522dda986`, the 0358 changed-route source
  used by 0362yd. Its existing probe uses the pre-patch `Edit::plan` API and
  therefore is not silently presented as a 0362yf patch-protocol endpoint.
- External fixture: `/home/zhuhe/code/litchi-worktrees/.coord/odp-native-qualification-fixture-0362yd.odp`,
  4,215,423 bytes, SHA-256
  `078a83cd132a308e4609cf46f587cb68c1f928c8d2e6e1002d344223ec385cb5`.
  Its generated `content.xml` is 591,185 bytes with SHA-256
  `4819e5022db80c5b3b0ee4ae23ea99932c0d703635ef613d17245d375ecfaa93`;
  the fixture follows the tracked 0362yd 2,000-slide/4-MiB-media recipe.

## Build boundary

The candidate probe compiled successfully in the external target
`/home/zhuhe/code/litchi-worktrees/.coord/odp-native-qualification-candidate-f28a599077`
with one job, release profile, debug information disabled, incremental
compilation disabled, `--offline --locked`, and the 20-GB capped cargo shim:

```text
cargo build --offline --locked --release \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin odp_native_qualification_probe
```

The successful diagnostic binary predates the final content-scope assertion
in `1b3ec77fa5`; it is retained only as external build evidence, not as the
final candidate endpoint. No baseline rebuild was started after the
coordinator released the build slot.
The historical 0362yd binaries remain external and untouched. No Callgrind or
native timing process ran.

## Untimed proof boundary

The diagnostic candidate execution reached these checks successfully:
source/no-op identity, semantic read, append change, fixture `content.xml`
source proof, changed append output, same-owner replay, content-matching
replay, changed-recipient refusal, and the scoped recipient's content/member
preservation checks. ZIP inspection showed that inverse replay restored the
exact `content.xml` bytes while the newly published content entry carried
different ZIP timestamp/attribute metadata. The reviewed receipt therefore
keeps the reopened inverse proof scoped to `content.xml` and separately proves
opaque-member preservation. The probe also requires a direct inverse applied
to the committed owner to restore the complete source archive byte-for-byte;
that whole-owner proof is distinct from the reopened publication proof.

The corrected source is committed in `1b3ec77fa5`; this follow-up adds the
whole-owner inverse proof and removes the quiet-host rejection from the
untimed `--correctness-only` path. A correctness-only run against two
compatible endpoints remains pending; strict quiet-host checks still apply to
calibration and ABBA timing, and no performance ratio is accepted by this
receipt.
