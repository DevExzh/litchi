# 0362ugjb — PPT lazy build receipt handoff

This packet materializes the final 472b PPT lazy qualification manifest and a
truthful `build.v1` receipt adapter.  It records already completed builds and
the accepted 28-route/4-fixture correctness and allocation preflight; it does
not claim a rebuild, native timing, calibration, allocator qualification, or
ABBA acceptance.

The external packet root is
`(deleted-coord-artifact)`:

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
  --manifest (deleted-coord-artifact) \
  --projection (deleted-coord-artifact) \
  --preflight-receipt (deleted-coord-artifact) \
  --logs-root (deleted-coord-artifact) \
  --output (deleted-coord-artifact)
```

The historical adapter used the frozen runner's manifest, probe-source, and
build receipt validators with the projection path.  The repaired adapter uses
the public runner lifecycle described below and retains the frozen identity as
an independent input.  Both rehash the baseline endpoint and every reviewed
overlay file.  Baseline admission is exact: HEAD,
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
The sparse baseline has roughly 450k index entries: an unrestricted porcelain
status took 15.65 seconds in the review environment and exceeded the earlier
10-second observation budget.  The runner and adapter now use Git's
read-only `--no-optional-locks` mode for status reads; it produced byte-for-byte
the same five-line status in 0.21 seconds.  This bounds the source command
without weakening the exact status, HEAD/tree, or overlay-content fences.

The supplemental cold first-commit route is documented separately in
`0362ugja-ppt-lazy-cold-first-commit.md`.  It uses CPU 29, `warmups=0`, and
one no-op iteration, but is not part of the frozen 28-route acceptance
manifest.  CPU 29 is reserved for a future quiet execution only.

## Repaired public-lifecycle packet

The initial 472b packet above is retained as historical evidence.  The
reviewable source-only handoff uses the separately materialized repaired
packet at
`(deleted-coord-artifact)`:

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `ppt-lazy-edit-472b-repaired.json` | 9829 | `8cdcc3644bf072664d37651923245df0ff1c5425a3f73e695db3a3b77ec5dc15` |
| `ppt-lazy-build-input-projection-472b-repaired.json` | 9724 | `762ba94ec10b83cb574d9d2b01a1992316a3a1592b31d42b8f6ca89d152250f6` |
| `ppt-lazy-build-receipt-472b-repaired.json` | 9976 | `bc44f06bf77959edf0a67a43ea0be15701d547dc5b490395a59601ab27641d45` |
| `ppt-lazy-build-adapter-472b-repaired-public-v3.json` | 3885 | `e8d5016528e58c2ee4e8e4c63c39c44fc88faa500ed81176c9e25c4d90b387b7` |

The repaired manifest binds normal and allocator targets to their actual
binary parents (`baseline-082-normal-target`, `baseline-082-allocator-target`,
`candidate-472b-normal-target`, and `candidate-472b-allocator-target`), rather
than to their shared parent.  Its baseline adapter includes the exact five
status lines and the ten-file content closure.  The public runner now accepts
the full manifest plus `--build-input-projection`, validates the canonical
projection digest, resolves both endpoints, rechecks locks/executables/target
directories/worktree state, and then validates the build receipt.  The adapter
uses that public `validate_inputs` lifecycle; it does not construct endpoint
objects or invoke frozen-runner internals.  It also rechecks the projection
identity at every later endpoint lifecycle boundary.  The public runner identity is
86075 bytes, SHA
`88f10be52bf02a8d5210a831d38af4358738ea35aea956341d7d4e8af1342977`; the
historical candidate runner remains independently pinned at 73596 bytes, SHA
`19d80d7f6e9450b60862929c4d864c0caae213aaa9e0682fca1c5fcc67d28c4f`.

The repaired adapter completed `status=structural_pass` after rehashing the
source archive and all endpoint/overlay identities; no Cargo, native probe,
calibration, allocator measurement, or ABBA run was performed.  The earlier
`ppt-lazy-build-adapter-472b-repaired.json` is retained as the pre-allowlist
receipt, and the earlier `-public` and `-public-v2` artifacts are retained as
intermediate public-lifecycle receipts.  The `-public-v3` artifact is the
final receipt bound to the public runner, exact four-path probe allowlist,
projection rechecks at use, and bounded read-only Git status.

Its exact source-only invocation was:

```sh
python3 tools/ppt_lazy_build_receipt_adapter.py \
  --manifest (deleted-coord-artifact) \
  --projection (deleted-coord-artifact) \
  --preflight-receipt (deleted-coord-artifact) \
  --logs-root (deleted-coord-artifact) \
  --output (deleted-coord-artifact)
```

## Integration validation

The coordinator's composed tree passes all 26 tests with
`PYTHONPATH=tools python3 -m unittest tools.test_ppt_lazy_build_receipt_adapter tools.test_ppt_lazy_cold_first_commit tools.test_ppt_lazy_native_qualification`
(0.174 seconds). The initial invocation without `PYTHONPATH=tools` passes
11 tests but fails to import the native qualification suite's bare sibling
module. Record 0362ul41 tracks fixing the current full tooling-row import
setup and registering the two new suites; this composed result does not
claim that gate row passed.

These source/tooling changes retain the recorded 472b build identities.
The subsequent PPT raw-parser API repair requires fresh candidate bindings
and validation before final native acceptance. The cold primary-noop route
remains supplemental and does not close the changed external-media workload.
