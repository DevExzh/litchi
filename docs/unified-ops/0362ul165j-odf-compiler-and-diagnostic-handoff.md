# 0362ul165j — ODF compiler and diagnostic handoff

The DrawingML test repair is merged into trunk as `11de659154`, following
its passing isolated leaf Clippy check. Its Q worktree remains an active
input to the remaining-checks diagnostic. The ODF tooling leaf remains
unmerged while concrete compilation and preparation defects are repaired.
The latest branch census still has no abandoned unified leaf to remove.

## Actual ODF compilation

The first ordinary-example check at `ef033324` stops under `--locked`:
the new ODF features require seven local packages absent from the standalone
performance lock. Root preserves the failed run, then resolves metadata
offline under a 20 GiB cap to repair the lock intentionally. Comparison of
the parsed lock graphs confirms that every existing registry package entry
is identical. Only seven local ODF packages and their `litchi-ops` edges are
added. The new lock has 57,961 bytes and SHA-256
`89ee807d671d102c215e9067cb19814f67c81794cd250c4a94253bac329506a4`.

Root also finds that the raw-codec producer builds debug binaries and has no
separate optimized producer. Such protocol output cannot close the required
release performance comparison. Leaf commit
`16744934d616378a7cc8533aa2ba93dd65513f00` migrates all five live
ODF/DrawingML/Signed lock consumers together, pins the release profile and
binary directory, and updates the producer, receipt consumer, probe
provenance and tests. Historical baseline identities remain unchanged.
Fifteen focused source tests pass; this is not native qualification.

A fresh guarded ordinary/allocator `cargo check --release --offline --locked`
diagnostic then resolves all dependencies and reaches the Rust probe. The
ordinary check fails with three errors: unresolved `zip` import, clearing a
buffer before its borrowed event is finished, and distinct `impl FnMut`
types in match arms. The allocator check is unrun. Source and lock hashes
remain unchanged and no resource failure occurs. The worker must fix these
with static dispatch and the existing measurement boundaries intact.

An additional preparation mismatch remains explicit: production C/Q has the
old performance lock, while the separate consumer harness requires the new
lock. The candidate preparation currently overlays only the manifest and
probe. It needs an exact, recorded consumer-lock overlay with unchanged
production source. Merging the probe into C would also violate the current
requirement that the candidate probe destination be absent. C and P stay
separate; a new receipt cannot relabel either source or bypass these checks.
The ODF raw row is returned, leaving 24 of 28 source preparations ready.

## Remaining-checks diagnostic

The proposed Q driver covers precisely the 24 unrun full53 rows 29–52 and
five runtime supplements. Root returns v1 before launch: it passes the rich
source-audit dictionary into `run_safe_gate.run_step`, whose interface requires
a distinct dictionary containing `cargo_locks`. Its six helper tests miss
that executable handoff.

V2 preserves both source identities separately. Thirteen controls exercise
the execution loop, ordinary-error continuation, resource/provenance stops,
OfficeCLI preservation before row 38, invalid statuses, nested process holders
and signal/OOM failures. A small Python child passes through the actual
approved runner with the actual Q source identity. Root also compares all
29 commands with the original plan after exact source/state/target path
rebinding. Preparation and dry-check pass; GO has not run at this checkpoint.
The driver and controls are committed under `results/0362ul165j/` with bytes
identical to the reviewed external copies.

This remains a diagnostic. It imports no passed rows from the failed C33
run, and cannot qualify the eventual complete all-65 gate. All prior drivers,
plans, terminal failures and review receipts remain preserved.

## Further source-review returns

W6 v4 is returned because the public recipe rejects its actual origin HEAD,
its expected probe hashes precede its own overlays, nested runner output
escapes the live bounds, and its consumer accepts extra routes. V5 must
exercise the actual origin/overlay/public handoff rather than relying only
on the eight passing source tests.

Flat-parts v3 repairs the prior findings but is returned for mismatched
baseline features/API, allocator instrumentation in baseline ordinary
measurements, insufficient report revalidation, unbound Cargo feature/command
identity and parent-symlink admission. V4 must preserve the original baseline
algorithm while fixing both endpoint roles and their real producer/consumer
paths. No restored workload is dropped or waived.

All 28 actual qualification groups remain open. C/P stay at C33 during these
repairs; successor production selection, consumer integration, final gates,
source rebinding and final control cleanup remain required. Only memo 22 is
deferred. Nothing is pushed.

## Evidence

Paths are relative to `/home/zhuhe/code/litchi-worktrees/.coord/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `odf-raw-example-compile-ef03-20261026/result.json` | 5,525 | `751562c0ee0b7dd4a3497a575fe95252de04eaf03558ddf1a968e770b52e7a5f` |
| `odf-raw-lock-repair-20261026/result.json` | 1,151 | `38f1471c7cc63cdae241a6ccb81ce56ed0cdce79fdec8127499bae797fc328ae` |
| `odf-raw-lock-repair-20261026/metadata.json` | 1,469,712 | `d0f678eae25c6473d14e0b03d5cdac88e44160addfdbeba41807f7edf22dbad1` |
| `odf-raw-example-compile-1674-20261026/root-return.json` | 1,917 | `512e1983ae8c80e0af4bf825a9de876f216a58787ef2099ab1e8480d91825ae9` |
| `odf-raw-example-compile-1674-20261026/result.json` | 5,624 | `3252c2225e91a6a74f2703c060384150cefd383705c3a085527a175fee871061` |
| `final-gate-q472a-rest29-diagnostic-20261026-v2-root-review/review.json` | 2,721 | `327f49448a77df5e02406f79364612b4a854d0f92513d5278edcea0e8fa7b87f` |
| `final-gate-q472a-rest29-diagnostic-20261026-v2-root-review/handoff-control.json` | 3,573 | `481c51a0d141317a23bf66ec834a1fc13c41bb7ed5a8a7e4d3035f1b9cf8f638` |
| `final-gate-q472a-rest29-diagnostic-20261026-v2-packet/plan.json` | 87,551 | `d8e3beb64dc7667d5a157ae04990ea917ba128e37289fc55236511cb9594ba8a` |
| `w6-odt-ods-C2a-prep-20261026-v4-independent-review-20261026/review-receipt.json` | 11,816 | `71861aec65a30f6ee632ff244d00b8c9cc1062f4ccb265cbfbe8fd576f0dbdd6` |
| `odf-flat-parts-C2a-prep-20261026-v3-independent-review-fresh-20261026/review-receipt.json` | 11,816 | `0230aba2969dd5012d10fec431ddc2bf8c0403108166d8eff2a2b9e05e8b4f06` |
| `odf-flat-parts-C2a-prep-20261026-v2-independent-addendum-20261026/addendum.json` | 4,777 | `561a55712e465e69358fdbfa4bad0f7f534ab1cfdfef9a1e2af54997490d6618` |

The external metadata artifact exceeds 1 MiB and is not committed. Its
regeneration command, run from the exact pre-repair ODF leaf with the preserved
input lock, is:

```sh
LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 \
  /home/zhuhe/code/litchi-worktrees/.coord/bin/cargo metadata \
  --offline --format-version 1 --manifest-path tools/perf-baseline/Cargo.toml
```

This command intentionally updates the consumer lock; it is not a validation
command for a frozen source. Subsequent compilation retains `--locked`.
