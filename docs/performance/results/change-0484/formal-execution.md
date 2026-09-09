# Formal route and axis execution

Attempt `formal1` freezes `route-protocol.json` after the diagnostic harness
checkpoint. It measures the three replay routes and the separate one-factor
input, sink, and compression inventory. This is enabler characterization,
not a historical before/after optimization comparison. The full non-iWork
performance goal remains open.

The normal and allocator release executables were built separately with
unchanged, identical 7,157-file source manifests. Both use source digest
`5a86d19865374a7e914ebb2f168273800b4e04de5ae287bd7ecc4fe37c29b0ae`.
The build records in `route-attempts/formal1/` bind the copied executable hashes,
build commands, environment, protocol, and gate receipts. These matched builds
supersede the differing diagnostic source scopes for formal measurements.

All 60 route and 54 axis pilots pass in this attempt. Their whole-lane gate
receipts have unchanged sources matching both builds. The subsequent
`execution-input-binding-formal1` gate also passes. Its output,
`route-attempts/formal1/execution-inputs.json`, binds exact prepared archive
hashes and their export receipt, build/protocol/machine hashes, a current
source snapshot, and observed device/mount records. The workload protocol
itself has no Rust-source snapshot field; source identity is established by
the matched builds, this execution snapshot, and the capture gates. The
analyzer must verify those linked records before accepting formal results.

Before freezing, the capture handlers were adjusted to retain a terminal failed
receipt for errors from the shared validator as well as route-specific errors.
`stream-formal-receipt-tests-dev110` passes all 23 Python tests, including a
regression that checks the failed receipt and every retained artifact hash for
both route and axis captures. No production source changed.

`prepare_axis_inputs.py` invokes the frozen normal executable to export the
64- and 131,072-paragraph source fixtures. This setup uses excluded diagnostic
operations and retains the original archives, candidates, manifests, and
reports under `axis-input-origin/formal1/`. It checks each archive against its
native manifest and independently regenerates the expected source XML hash in
Python. The three file-input paths under `axis-input/` are exact exclusive
copies of those verified sources. The two 64-paragraph paths deliberately
contain identical source bytes; their authored workloads differ.

The source archives are 2,106 and 343,935 bytes. The source-file preparation,
fingerprinting, and fixture/report owners remain outside sample timing. Source
and replay files use the recorded workspace filesystem; workspace and cache
scratch directories resolve to the same device on this host. Cache eviction is
not performed. File replay uses the explicit data-sync policy; these runs do
not measure atomic save or cold-cache behavior.

The reproducible command sequence from the repository root is:

```sh
python3 -B docs/performance/results/change-0484/measure_routes.py freeze
python3 -B docs/performance/results/change-0484/measure_routes.py build-normal --attempt formal1
python3 -B docs/performance/results/change-0484/measure_routes.py build-allocator --attempt formal1
python3 -B docs/performance/results/change-0484/gate.py --attempt formal1 axis-input-preparation python3 -B docs/performance/results/change-0484/prepare_axis_inputs.py --attempt formal1
python3 -B docs/performance/results/change-0484/gate.py --attempt formal1 route-pilots python3 -u -B -c 'import sys; sys.path.insert(0,"docs/performance/results/change-0484"); import measure_routes as m; m.run_all("formal1",pilot=True)'
python3 -B docs/performance/results/change-0484/gate.py --attempt formal1 axis-pilots python3 -u -B -c 'import sys; sys.path.insert(0,"docs/performance/results/change-0484"); import measure_routes as m; m.run_axis_all("formal1",pilot=True)'
python3 -B docs/performance/results/change-0484/gate.py --attempt formal1 execution-input-binding python3 -B docs/performance/results/change-0484/bind_execution_inputs.py --attempt formal1
python3 -B docs/performance/results/change-0484/gate.py --attempt formal1 route-captures python3 -u -B -c 'import sys; sys.path.insert(0,"docs/performance/results/change-0484"); import measure_routes as m; m.run_all("formal1",pilot=False)'
python3 -B docs/performance/results/change-0484/gate.py --attempt formal1 axis-captures python3 -u -B -c 'import sys; sys.path.insert(0,"docs/performance/results/change-0484"); import measure_routes as m; m.run_axis_all("formal1",pilot=False)'
```

Freeze and attempt receipts refuse replacement. Reproduction must use a fresh
checkout/output location and attempt token rather than overwrite retained
evidence. The gate holds the shared CPU lock and snapshots sources before and
after each lane; its child invokes the library entrypoint to avoid acquiring
the same lock twice. The standalone measurement CLI holds the lock itself and
must not be nested inside `gate.py`.

The pilot inventory has 60 route and 54 axis processes, with three measured
operations and one warmup each. Pilots are excluded from formal statistics.
The formal inventory has 120 route and 108 axis processes, with 30 samples and
three warmups each. Repeat two reverses the inventory order. A successful
whole-lane gate and every child receipt are both required for acceptance.
Profiling runs are separate from these captures. Incomplete or failed
inventories cannot support a completed measurement claim.

## Completed capture and analysis gates

Both formal lane gates pass with the matched source digest above: 120 route
and 108 axis reports, all 6,840 measured samples accepted. The 24 separate
external profiles in `route-profiles/profiles1/` also pass with unchanged
sources. Their raw data covers whole children and remains separate from
formal timing statistics.

The `evidence-tests-formal1` gate passes 48 Python tests. The
`route-analysis-verification-formal1` gate independently recomputes and matches
`route-summary.json`, including the canonical frozen per-report validators,
cross-route/role/repeat corpus identity, exact input and build custody, and
separate metric scopes. `route-render-formal1` produces the complete table,
CSV and scaling figure from that verified summary.

`replay-directory-cleanup-formal1` removes exactly 60 empty, accepted file-route
directories. `cleanup-replay-dirs-formal1.json` binds each removed directory to
its accepted child receipt. The analyzer accepts an absent replay directory
only through that cleanup evidence; it still rejects symlinks and nonempty
directories. Exported corpus archives remain retained inputs, not scratch.

The six retained `perf.data` captures were exported under the
`route-perf-exports-exports1` gate using `export_route_perf.py`. The profile
summary is regenerated after export so each text stack has a source binding.
The separate `route-input-metadata-metadata2` gate uses
`profile_input_metadata.py --attempt metadata2 --build-attempt formal1
--raw-authored` with the normal copied executable; it adds four syscall
summaries and two authored-heavy raw traces, each with one sample and one
warmup. These diagnostic children are excluded from the 6,840 formal samples.

The initial metadata1 attempt failed before benchmark launch because strace
rejects `pread` as a syscall name. Metadata2 uses the actual `pread64` name;
the earlier failure is retained. `metadata_evidence.py compact` validates
the raw hashes against lossless gzip decompression before removing the two
redundant raw files. `metadata_evidence.py verify` rechecks that custody.

Final evidence validation: `evidence-tests-final1` passes all 53 Python
tests with unchanged source. `owned-temporary-cleanup-final1.json` records
generated cache removal. The final seal verifies all 228 formal processes,
24 route profiles, six metadata profiles and lossless compressed trace custody.
