# 0957 bounded allocator evidence packet

The external 0957 capture completed all 27 planned children with zero failures,
and the one-child tiny preflight passed. This repository packet now contains
the final bounded plan, sealed packet manifest/state, current allocation
summary, final v2 tooling, control wrappers, and the source/build/quality
readiness receipts. The summary is descriptive allocator evidence only; it
does not make a speedup, adoption, or performance decision.

`final-packet-provenance-v1.json` binds the copied files to their external
paths, sizes, and SHA-256 values. Raw reports/logs, source manifests,
binaries, and full reducer outputs remain at the external packet root
`/home/zhuhe/litchi-performance-0957`. The external source and allocator
binary are referenced by descriptor only, in accordance with repository size
and binary rules.

The root-owned final report is now available at
`../../0957-xlsx-managed-compression-allocations.md`, and the bounded final
allocation review, findings, summary helper, and reducer amendment receipt are
copied here. The report retains the no-adoption decision and its threshold and
phase caveats. Cleanup removed 2,949 temporary files totaling 2,674,731,403
logical bytes; see `receipts/cleanup-v2.json`. The earlier
pre-capture records remain available as snapshots: `README-v1.md`, `README-v2.md`,
`source-build-provenance-v1.json`, and `source-build-provenance-v2.json`.
Those snapshots are not rewritten by the final bounded packet descriptor.

Regenerate source/build artifacts using `source-build-provenance-v2.json`.
The external source manifest is 2,267,268 bytes, SHA-256
`f5fd1ee75b734bb7efaa1882be0397defea98c76c835be0bdf1f565c573ef617`;
its exact reproduction script is `tools/reproduce_source_manifest_0957_v1.py`.
The retained allocator binary is 43,676,872 bytes, SHA-256
`9ba37a6b9796ad0e316e7571ebe117ba7a5f97bcf5d29153c7cb3d74c485e9a5`.
Neither belongs in Git.

With the retained external source/binary in place, the capture regeneration
command is (use a fresh destination):

```sh
python3 -B /home/zhuhe/litchi-performance-0957/tools-v2/capture_0957_v2.py --plan /home/zhuhe/litchi-performance-0957/plan-final-v2.json --packet /home/zhuhe/litchi-performance-0957/capture-reproduction-new
```

The sealed capture has 194 manifest-listed files totaling 2,129,250 bytes;
`capture-allocation-v1/packet-manifest.json` supplies every file name, size and
SHA-256. Recompute its descriptive summary without rerunning workloads:

```sh
python3 -B /home/zhuhe/litchi-performance-0957/tools-v2/reduce_0957_v2.py --plan /home/zhuhe/litchi-performance-0957/plan-final-v2.json --packet /home/zhuhe/litchi-performance-0957/capture-allocation-v1 --output /home/zhuhe/litchi-performance-0957/allocation-summary-reproduction-new.json
```

The original plan binds the pre-capture reducer. Its post-capture import and
receipt compatibility corrections are explicitly bound by
`receipts/reduction-amendment-v1.json`; the raw packet and statistical contract
are unchanged. `validation.json` records the final repository checks.
