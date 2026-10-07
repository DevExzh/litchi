# 0949 reduced allocation-attribution evidence

The [change record](../../0949-xlsx-multi-window-allocation-attribution.md)
explains the measured result and limits. This packet contains only bounded
summaries, scripts, plans, and receipts. Raw profiles, decoded trace aggregates,
and full exports remain outside the repository. No benchmark binary is committed.

## Authoritative artifacts

- `heaptrack-plan-v3.json` and `tooling/run_heaptrack_0949_v3.py`: frozen
  four-child schedule and inert-until-capture runner.
- `receipts/capture-artifact-index-v1.json`: input files hashed before decoding.
- `receipts/0948-reuse-reverified-v1.json`, capture preflight/postflight:
  source/build/binary bindings for reused 0948 executables.
- `raw-decoder-v1/`: complete-process summary and terminal receipts.
- `attribution-final-v1/`: final classification, using only
  `tools/attribution-config-root-v1.json` and the helper hash bound in
  `receipts/root-offline-attribution-seal-v1.json`.
- `review/report-admission-v2/result-final-v2.json` and
  `review/heaptrack-conservation-v1.json`: independent admission and conservation.
- `review/xml-multiwindow-source-profile-reconciliation-v2.json` and its
  `.py` / `.md` siblings: exact dual-frame call-site selection, disjoint source
  contributions, source hashes, and corrected residuals.

The external packet also preserves draft plans, excluded mixed-config
`attribution-v1`, superseded author `attribution-v2`, the recovered original
classification config, and the export-helper failure before subprocess launch.
They are provenance, not alternate admitted results. `allocation_calls` already
includes reallocations; do not add `reallocation_calls` to it.

## Offline regeneration

Set `packet` to an extracted full 0949 packet and `source` to the preserved
0948 `candidate-build-v3` tree, recovered from the dependency archive if needed.
Use a new output file; the reconciler is create-only:

```sh
python3 -B "$packet/review/xml-multiwindow-source-profile-reconciliation-v2.py" \
  --packet "$packet" --source-root "$source" --output /tmp/0949-reconciliation.json
```

The runner's `--validate-only` checks original source/binary/tool bindings
without a workload. Its `--capture --go` command is the recorded way to regenerate
fresh diagnostic measurements under the frozen plan; it requires a fresh capture
destination and the original dependency layout. Offline replay of archived raw
captures is the appropriate way to reproduce these exact numerical results.

## External archive

`/home/zhuhe/litchi-performance-0949-archive/archive-0949-v1.tar.zst`
contains the full 202-file packet (717,132,537 source bytes).

- Archive: 10,206,299 bytes, SHA-256
  `d38a2d38904de6d49cf186326332268da3199734be3215ba12b65ace90fb3a54`.
- Manifest: 49,561 bytes, SHA-256
  `58070fbcc5ee8f2da8f7edf3a8d369f386ef3715b349c3710e59b3019c7b1b56`.
- External dependency: the 0948 archive, 1,348,106,492 bytes, SHA-256
  `15a4008a6b888fed4fea50f0746d9787df32e1a2df6b164d25a7dabb19f1db95`.
  Its source/build/binary identities are recorded in `archive/dependency-0949-v1.json`.

The archive helper and complete manifest are included under `archive/`.
Regenerate the archive with `archive-0949-v1.py --seal --go --packet FULL_PACKET
--archive-root NEW_ARCHIVE_DIRECTORY`, using the recorded dependency layout.
The helper records the selected file sizes, modes, and hashes; compression
container bytes may depend on filesystem metadata. An initial tar option-order
failure occurred before compression and is retained in the attempt receipt.


## Relocated replay and cleanup

The extracted archive passed the complete file census, size, mode, and SHA-256
check. Offline replay ran the decoder, attributor, and source-site reconciler;
the reconciler consumed the freshly decoded aggregates. The comparison-only v2
validator reports 36 exact file comparisons, 9 normalized JSON comparisons,
and 12 fresh-decoder hard-link checks, all passing. Only location fields and
descriptors of location-bearing summaries are normalized; allocation counts,
requested bytes, site classes, residuals, and underlying aggregate hashes match.
The explicit 0948 source dependency remains external.

The initial replay completed all analysis stages but its final comparator did
not account for moved paths. The failure receipt is retained. The corrected
comparison reused those successful outputs; it did not rerun a workload or
profiler. The replay helper and comparison-only validator are included under
`archive/`; invoke their `--help` for the packet/source/output arguments.
The final archive, canonical packet, and dependency sources/binaries are retained.
Cleanup receipts identify only this batch's removed extraction, replay, and
Python-cache scratch.
