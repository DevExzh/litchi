# 0362yf — e4a baseline probe adapter

The native qualification candidate probe at
`tools/perf-baseline/src/bin/odp_native_qualification_probe.rs` targets the
post-0358 ODP API (`edit_tail_append().commit()` and
`apply_tail_append(&patch)`). The pinned original baseline
`e4a0614f5d3a1d768346cc33a4a3b86522dda986` predates that surface. It must be
built with its own source-backed publication lifecycle; compiling the
candidate probe against the old checkout would silently change the baseline
workload.

## Static adapter and build selection

`tools/perf-baseline/odp_native_qualification_e4a_probe.rs` is the protocol
adapter. It is copied into the old checkout's untracked
`crates/litchi-odp/examples/` directory only for the external build. It uses
the pinned checkout's pre-Patch `SourceBackedPresentation`,
`tail_append::Edit::plan`, and materialize/write APIs while consuming the same
external frozen fixture and `--mode/--fixture/--route` arguments as the current
probe. The older
`tools/perf-baseline/odp_native_qualification_e4a_route.rs` include wrapper and
`docs/unified-ops/results/0362ye/odp_0362ye_probe.rs` remain historical
diagnostic sources; they are not the matched runner protocol. The baseline
checkout and its protected target directory are not changed by this commit.

The pinned source identities are:

| source | identity |
|---|---|
| e4a commit | `e4a0614f5d3a1d768346cc33a4a3b86522dda986` (revision) |
| external-fixture protocol adapter | `47cc7a780d93b23015ff698216e98ce9597ee300ee1b34ae21a39276b2baded7` |
| historical route adapter | `7c79858b4626135f4b7cdb695fca5e185d76665e5717c1ec7e610b3ec3a03875` |
| included e4a route | `fc08a3d65ec38c9053b1927e25943762c06c6894026e186b011bfb0dab028bb9` |
| common probe source | `9d709068b90f685fd45b367320296484020c4af45bd42da32d7ac13d5ac41650` |

At the next authorized build slot, select the adapter explicitly. The copied
example is a temporary generated overlay, not part of the baseline source
identity. Keep the build receipt and all binaries outside both worktrees:

```text
set -eu
ADAPTER_SRC=/path/to/unified-ops-api/tools/perf-baseline/odp_native_qualification_e4a_probe.rs
BASELINE=/path/to/uo2-odp-heap-repair-0362yd-baseline
ADAPTER_DST=$BASELINE/crates/litchi-odp/examples/odp_native_qualification_legacy_probe.rs
TARGET=/external/odp-0362yd-baseline-target
RECEIPT=/external/odp-qualification/0362yf-baseline-build-receipt.txt
ADAPTER_SHA=47cc7a780d93b23015ff698216e98ce9597ee300ee1b34ae21a39276b2baded7
test -z "$(git -C "$BASELINE" status --porcelain --untracked-files=all)"
test ! -e "$ADAPTER_DST"
cleanup() {
  rm -f -- "$ADAPTER_DST"
  test -z "$(git -C "$BASELINE" status --porcelain --untracked-files=all)"
}
trap cleanup EXIT
cp -- "$ADAPTER_SRC" "$ADAPTER_DST"
test "$(sha256sum "$ADAPTER_SRC" | cut -d' ' -f1)" = "$ADAPTER_SHA"
test "$(sha256sum "$ADAPTER_DST" | cut -d' ' -f1)" = "$ADAPTER_SHA"
mkdir -p "$(dirname "$RECEIPT")"
{
  printf 'baseline_head '; git -C "$BASELINE" rev-parse HEAD
  printf 'adapter_source_sha256 '; sha256sum "$ADAPTER_SRC" | cut -d' ' -f1
  printf 'adapter_overlay_sha256 '; sha256sum "$ADAPTER_DST" | cut -d' ' -f1
  printf 'target_dir %s\n' "$TARGET"
  printf 'command cargo build --offline --locked --release --manifest-path %s -p litchi-odp --example odp_native_qualification_legacy_probe\n' "$BASELINE/Cargo.toml"
} >"$RECEIPT"
CARGO_TARGET_DIR="$TARGET" cargo build --offline --locked --release \
  --manifest-path "$BASELINE/Cargo.toml" \
  -p litchi-odp --example odp_native_qualification_legacy_probe
BINARY=$TARGET/release/examples/odp_native_qualification_legacy_probe
test -x "$BINARY"
printf 'binary_sha256 '; sha256sum "$BINARY" | cut -d' ' -f1 >>"$RECEIPT"
printf 'binary_bytes '; stat -c '%s' "$BINARY" >>"$RECEIPT"
test "$(sha256sum "$ADAPTER_DST" | cut -d' ' -f1)" = "$ADAPTER_SHA"
# EXIT trap removes only the generated overlay. The receipt and binary remain.
```

The build receipt must preserve the resolved e4a `HEAD`, the source and copied
overlay adapter SHA-256 values, exact command, binary byte count/SHA-256 and
external target path. The final overlay hash check happens before the `EXIT`
trap removes only the named generated example. Verify
`git -C "$BASELINE" status --porcelain --untracked-files=all` is empty after
the trap; do not remove or reset any other file. A separate clean checkout at
the same e4a commit may be used instead, but its copied source and resulting
binary must carry the same receipt identities. The qualification manifest may
then resolve the clean baseline `HEAD` and binary without weakening its source
cleanliness check. The protocol adapter's routes are exactly `open`, `read`,
`write` and `append`; it has no `replay` route. The old historical probe's ten
positional routes remain available only to the legacy allocator/Callgrind
diagnostics and are not used for matched native timing.

## Proof boundary

The e4a adapter preserves these baseline semantics:

| Capability | e4a adapter | current candidate probe |
|---|---:|---:|
| source-backed open/read and unchanged materialize/write | yes | yes |
| bounded append publication and source/target content proofs | yes, via `Edit::plan` | yes, via `Commit` |
| same-owner `Patch` replay | no | yes |
| equal-content cross-copy `Patch` replay | no | yes |
| direct whole-owner inverse proof | no | yes |
| reopened inverse content proof | no | yes |
| changed-recipient patch refusal and opaque-member preservation | no | yes |

The e4a adapter's `independent_append` output deliberately repeats an
independent old publication plan. It is a deterministic correctness check, not
patch replay, and is not a timed route. The old endpoint cannot prove the
candidate's patch base matching, inverse, or changed-recipient behavior. The
current candidate's correctness preflight remains strict for every one of
those newer checks; this adapter is only the compatible original baseline for
the matched timed routes.

No Cargo build, timing run, or protected-artifact mutation was performed while
preparing this adapter.
