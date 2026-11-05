# 0362yg — ODP native qualification build boundary

This is a reduced execution receipt for the reviewed 0362yf harness. It is
not a native timing qualification and does not replace the accepted 0362yd
baseline evidence.

## Frozen inputs

- Candidate source provenance: `f28a599077c8537c2de729a5c5cb2a4e1f85ec3d`,
  with the probe-only fixes now recorded in `1b3ec77fa5`.
- Historical baseline provenance: e4a
  `e4a0614f5d3a1d768346cc33a4a3b86522dda986`, the 0358 changed-route source
  used by 0362yd. The matched baseline was built from a separate clean
  checkout at that exact commit; its protocol adapter uses the pre-Patch
  `Edit::plan` API and is not silently presented as a 0362yf patch-protocol
  endpoint.
- External fixture: `(deleted-coord-artifact)/odp-native-qualification-fixture-0362yd.odp`,
  4,215,423 bytes, SHA-256
  `078a83cd132a308e4609cf46f587cb68c1f928c8d2e6e1002d344223ec385cb5`.
  Its generated `content.xml` is 591,185 bytes with SHA-256
  `4819e5022db80c5b3b0ee4ae23ea99932c0d703635ef613d17245d375ecfaa93`;
  the fixture follows the tracked 0362yd 2,000-slide/4-MiB-media recipe.

## Build boundary

The candidate probe compiled successfully in the external target
`(deleted-coord-artifact)`
with one job, release profile, debug information disabled, incremental
compilation disabled, `--offline --locked`, and the 20-GB capped cargo shim:

```text
cargo build --offline --locked --release \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin odp_native_qualification_probe
```

The successful diagnostic candidate binary predates the final content-scope
assertion in `1b3ec77fa5`; it remains external build evidence only, not the
final candidate endpoint. The original dirty e4a worktree already contained
an older ignored example, so the baseline build used the clean detached
checkout `(deleted-coord-artifact)`.
Its generated lock file was carried from the original e4a checkout under the
same source commit and had SHA-256
`04665b965a919412e29c8772956ddf6267264fa40e3e0ef7c9967fd2eb584171`.

The first `--locked` attempt failed before compilation because that clean
checkout had no lock file. The copied lock resolved that operational failure;
the first locked compilation of the original adapter then failed with exit
101: the `concat!` format string attempted named captures, `Args::parse` and
`run` had different error types, and the source had one partial move and one
unused `mut`. No binary was accepted from that attempt. Those diagnostics were
fixed in `61e117fe4d`; the resulting adapter SHA-256 is
`47cc7a780d93b23015ff698216e98ce9597ee300ee1b34ae21a39276b2baded7`. The
exact successful baseline build used that temporary adapter overlay, one
job, release debug information disabled, incremental compilation disabled,
offline/locked mode and the 20-GB capped cargo shim:

```text
CARGO_TARGET_DIR=(deleted-coord-artifact) \
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 \
LITCHI_MEM_MAX=20G \
cargo build --offline --locked --release --jobs 1 \
  --manifest-path (deleted-coord-artifact) \
  -p litchi-odp --example odp_native_qualification_legacy_probe
```

The resulting external binary is
`(deleted-coord-artifact)`,
2,241,664 bytes, SHA-256
`fb12f47974f5a99f6a3bc136b585d2218504efe51e86d46a03771c014cacbb56`.
The target directory is 316,582,738 bytes. After the build, only the named
generated adapter example was moved out of the clean checkout; the source
worktree status was empty. The moved overlay is retained externally at
`(deleted-coord-artifact)`
with the same adapter hash. The historical 0362yd binaries remain external
and untouched.

## Untimed proof boundary

The baseline-only verification used the frozen fixture and no timing mode:

```text
timeout 120 taskset --cpu-list 3 \
  (deleted-coord-artifact) \
  --mode verify \
  --fixture (deleted-coord-artifact) \
  --title 'Qualification appended title' \
  --body 'Qualification appended body'
```

The probe returned schema
`litchi.odp.native-qualification.legacy-probe.v1` and all six legacy checks
passed: exact no-op source, semantic read, changed append plan, source
`content.xml` proof, changed append output and independent-plan repeat. The
capability receipt explicitly reported `patch_replay=false`,
`inverse_replay=false`, `scoped_replay=false`, shared timed routes
`open/read/write/append`, and replay scope `independent_plan_not_patch`.
The runner's `validate_legacy_correctness_report` accepted the same receipt.

Observed proof identities were:

| item | bytes | SHA-256 |
|---|---:|---|
| fixture | 4,215,423 | `078a83cd132a308e4609cf46f587cb68c1f928c8d2e6e1002d344223ec385cb5` |
| fixture `content.xml` | 591,185 | `4819e5022db80c5b3b0ee4ae23ea99932c0d703635ef613d17245d375ecfaa93` |
| no-op output | 4,215,423 | `078a83cd132a308e4609cf46f587cb68c1f928c8d2e6e1002d344223ec385cb5` |
| append output | 4,215,657 | `6da5a24d8e711729f86a3c0734a24775c9c7677a58e5ce74ca9278770bd0983f` |
| independent append | 4,215,657 | `6da5a24d8e711729f86a3c0734a24775c9c7677a58e5ce74ca9278770bd0983f` |

The plan source/target content hashes were
`4819e5022db80c5b3b0ee4ae23ea99932c0d703635ef613d17245d375ecfaa93` and
`b07a3bddf627fd3a218245f7b3a6305fd717dbcc33aed3b42484d58fe3fffd71`.
The semantic read reported 2,000 slides, a present first slide, 105,778 text
bytes and text SHA-256
`13573ed9cbe078a4663767f3796f8ba8a2d4204a20b625de54f5bcd620c08189`.

No candidate rebuild, calibration, ABBA timing, allocator capture, or
Callgrind process ran in this receipt.

The earlier diagnostic candidate execution reached these checks successfully:
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

The earlier candidate correctness fixes are recorded in `1b3ec77fa5` and
`d99034c17e`; this receipt adds the clean e4a baseline build and legacy
protocol proof. A correctness-only run against two compatible endpoints
remains pending; strict quiet-host checks still apply to calibration and ABBA
timing, and no performance ratio is accepted by this receipt.
