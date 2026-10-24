# 0362ugjb — PPT lazy build receipt handoff

This packet materializes the final 472b PPT lazy qualification manifest and a
truthful `build.v1` receipt adapter.  It records already completed builds and
the accepted 28-route/4-fixture correctness and allocation preflight; it does
not claim a rebuild, native timing, calibration, allocator qualification, or
ABBA acceptance.

The external packet root is
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-20261024`:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `ppt-lazy-edit-472b-final.json` | 9434 | `4710ef25d9d6a919903df480264592882859f71d2f5bbb5e3aef64fc5222756b` |
| `ppt-lazy-build-input-projection-472b.json` | 9329 | `06f79d12fcdc62af1d50b051c817fa69aafc69994b999fdd15651ea3c2f0e479` |
| `ppt-lazy-build-receipt-472b.json` | 9840 | `68bdfc491ba9465b536ac14ae5f0fc6465ba8e2bdfd8f1a7bc755a67dced1b26` |
| `ppt-lazy-build-adapter-472b.json` | 3593 | `6e7dc5ac3477380440333952275dacb8202c3bf681debd2f5640f29640460736` |

The projection is a canonical copy of the final manifest with only
`build.receipt.bytes` and `build.receipt.sha256` removed.  The receipt path is
retained, and all source, lock, fixture, probe, command, feature, toolchain,
target, binary, and route fields remain bound.  The receipt binds its
`manifest` and `manifest_sha256` to this explicitly named canonical
build-input projection, not to the full final manifest; the projection is not
represented as the full manifest hash.  This is the narrow adapter for the
receipt/manifest identity cycle, not a provenance bypass.

The candidate endpoint is source `472b6705475b073e7fe08f6e11ee17fb1e29f838`
(tree `ebc1cee05f124c0c63977aff07a23288c19e7d00`) with lock SHA
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`.  The
historical endpoint is `8779d059add6c90c74552801372c44f770db9f0f` (tree
`27e0b56e2f8e3ab11b00c119604b5636f358dfe2`) with the pinned baseline lock
SHA `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184`.
The source archives and all four binary identities remain in the external
receipt; the candidate source archive is 6,634,147,840 bytes, SHA
`44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274`.

Run the source-only handoff with no Cargo or probe execution:

```sh
python3 tools/ppt_lazy_build_receipt_adapter.py \
  --manifest /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-20261024/ppt-lazy-edit-472b-final.json \
  --projection /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-20261024/ppt-lazy-build-input-projection-472b.json \
  --preflight-receipt /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/qualification-receipt-472b670547-corrected.json \
  --logs-root /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs \
  --output /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-final-472b-20261024/ppt-lazy-build-adapter-472b.json
```

The adapter reuses the frozen runner's manifest, probe-source, and build
receipt validators with the projection path, then rehashes the baseline
endpoint and every reviewed overlay file.  Baseline admission is exact: HEAD,
tree, and the following status are required (no blanket dirty-root waiver):

```text
M tools/perf-baseline/src/lib.rs
?? test-data/unified-ops/ppt-lazy-native-fixture.ppt
?? tools/perf-baseline/src/bin/ppt_lazy_native.rs
?? tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs
?? tools/perf-baseline/src/ppt_lazy_native.rs
```

It also checks the registration, fixture, three probe files, and five compile
fixtures against both endpoints and the preparation receipt.  The corrected
preflight receipt is 7646 bytes, SHA
`fbe59be89acb3a588b659c217f05ddf3c6ac49cdbeee2dce03c821e469714db8`; the two
recorded candidate build logs are each 39797 bytes, with SHA
`ca7297042ecfc6cd212126ba8fe92eff78cd023cd4d4f5d658ebb0cc72fdd223` (normal)
and `5bca2dad0e894e48f8a55d7263f8d6f99147c6b686fd0f4e8f1b351e1754cea0`
(allocator).

The command above completed with `status=structural_pass`; its output packet
is the adapter artifact in the table.  This source-only pass rehashed the
6,634,147,840-byte candidate archive and checked the frozen runner, preparation
receipt, corrected preflight receipt, build logs, locks, and binary identities.

The supplemental cold first-commit route is documented separately in
`0362ugja-ppt-lazy-cold-first-commit.md`.  It uses CPU 29, `warmups=0`, and
one no-op iteration, but is not part of the frozen 28-route acceptance
manifest.  CPU 29 is reserved for a future quiet execution only.
