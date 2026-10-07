# 0961 bounded evidence packet

The [change record](../../0961-xlsx-edit-staging-capacity-trial.md) owns the
interpretation and verification scope. Sources, raw benchmark reports, logs,
source manifests, build caches and compiled executables are external under
`/home/zhuhe/litchi-performance-0961/`. No compiled executable is committed.

`candidate.patch` records the isolated trial; it is not a production change.
`source-manifests.json` binds both full source manifests by size and SHA-256
and records source regeneration. Source extraction uses commit
`4e69e7cda66ad9691a3c8a4babe709a5c1145d39`, the pinned root lock overlay and
preserved file modes; the benchmark itself uses its separate pinned
`tools/perf-baseline/Cargo.lock`.

Quality receipts explicitly distinguish the full 1,403-test run from the final
34-test focused run after test-only strengthening. The initial compile failure
was a missing test import, not an intentionally failing control. The external
failed log remains part of the record.

The runner/reducer are reused from 0960 with their original protocol schema;
0961's plan binds new sources, candidate executables and the stricter screening
selector. Schema reuse is not sample reuse. The final plan and root GO receipt
are immutable capture provenance; never rerun capture against retained output
paths. A new trial requires fresh output paths and source/build bindings.

Reproduce reduction and selection using fresh output files:

```sh
python3 -B /home/zhuhe/litchi-performance-0960/tools/reduce_0960_v3.py \
  --plan /home/zhuhe/litchi-performance-0961/protocol/plan-final-v1.json \
  --native-packet /home/zhuhe/litchi-performance-0961/capture-native-v1 \
  --allocator-packet /home/zhuhe/litchi-performance-0961/capture-allocator-v1 \
  --output /tmp/litchi-0961-reduction-replay.json
python3 -B tools/select_screen_0961_v1.py \
  --report /tmp/litchi-0961-reduction-replay.json \
  --output /tmp/litchi-0961-selection-replay.json
```

The second command is relative to this packet directory. The independent raw
audit helper also runs from here with a fresh output path:

```sh
python3 -B tools/audit_raw_0961_v2.py /home/zhuhe/litchi-performance-0961 \
  --output /tmp/litchi-0961-independent-audit-replay.json
python3 -B tools/build_external_manifest_v1.py \
  --root /home/zhuhe/litchi-performance-0961 \
  --output /tmp/litchi-0961-external-manifest-replay.json
```

The inventory excludes the two source trees (their complete manifests are
bound separately), disposable build targets and Python bytecode caches.
Regeneration timestamps and command/output-path descriptors can differ;
artifact hashes and measured arithmetic must agree.

The final result is `reject_screen`. All 54 children validated and the
independent raw audit found zero mismatches. No production change or qualified
effect is claimed. The rejected candidate source remains available as evidence.

`external-artifact-manifest.json` inventories 545 retained files totaling
98,592,050 bytes outside the source trees. It records each file size and
SHA-256, including raw packets, full source manifests, logs and executables.
The two removed build caches are described by `receipts/cleanup-v1.json`;
historical build-output paths intentionally no longer exist. Regenerate the
executables with the retained build command/helper against the pinned source
and lockfiles, using a fresh target and receipt directory.

The hash-bound patch retains its six blank unified-diff context lines.
Whitespace checking excludes that evidence patch; all authored documentation
and helpers pass. Every committed evidence file is below 1 MiB.
