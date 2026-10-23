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

`tools/perf-baseline/odp_native_qualification_e4a_route.rs` is the build-time
adapter. It includes the reviewed e4a route implementation in
`docs/unified-ops/results/0362ye/route_before.rs`, whose mode identifies the
pin as `0358a-pre-change-e4a0614f5d`. The common probe source is retained at
`docs/unified-ops/results/0362ye/odp_0362ye_probe.rs`; it is copied into the
old checkout's untracked `crates/litchi-odp/examples/` directory only for the
external build. The baseline checkout and its protected target directory are
not changed by this commit.

The pinned source identities are:

| source | identity |
|---|---|
| e4a commit | `e4a0614f5d3a1d768346cc33a4a3b86522dda986` (revision) |
| adapter | `7c79858b4626135f4b7cdb695fca5e185d76665e5717c1ec7e610b3ec3a03875` |
| included e4a route | `fc08a3d65ec38c9053b1927e25943762c06c6894026e186b011bfb0dab028bb9` |
| common probe source | `9d709068b90f685fd45b367320296484020c4af45bd42da32d7ac13d5ac41650` |

At the next authorized build slot, select the adapter explicitly:

```text
ODP_PROBE_ROUTE=/path/to/unified-ops-api/tools/perf-baseline/odp_native_qualification_e4a_route.rs \
CARGO_TARGET_DIR=/external/odp-0362yd-baseline-target \
cargo build --offline --locked --release \
  --manifest-path /path/to/uo2-odp-heap-repair-0362yd-baseline/Cargo.toml \
  -p litchi-odp --example odp_0362ye_probe
```

The build must record the resolved e4a `HEAD`, adapter/source SHA-256 values,
binary SHA-256 and external target path. The old probe's routes remain
`open`, `first_slide`, `repeat_slide`, `first_text`, `repeat_text`,
`write_text`, `repeat_write_text`, `unchanged_write`,
`append_commit_write`, and `replay`; allocator and Callgrind route runners
continue to invoke those exact names. No fixture generator or route is
replaced by the current candidate probe.

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

The e4a `replay` route deliberately repeats an independent old publication
plan. Its output can be compared with `append_commit_write` for the historical
route workload, but it is not patch replay and must not be labeled as such.
Likewise, the old route can prove that two independently planned appends have
the same output shape; it cannot prove the candidate's patch base matching,
inverse, or changed-recipient behavior. The current candidate's correctness
preflight remains strict for every one of those newer checks; this adapter is
only the compatible original baseline for the matched timed routes.

No Cargo build, timing run, or protected-artifact mutation was performed while
preparing this adapter.
