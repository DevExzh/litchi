# 0362ul136 — signed-native v2 consumer boundary

This source-only follow-up completes the v2 preparation and execution
consumer for policy 0362ul74.  It does not freeze a candidate source: the
checked-in `CANDIDATE_BINDING` remains `None` until the coordinator supplies
the immutable post-134a production capture.  The pending path therefore still
fails before creating an output directory and does not relabel any 0362ul24
receipt.

## External consumer

`tools/native_qualification/signed_native_v2_consumer.py` is the only v2
consumer.  Once the policy binding is present, `prepare_signed_native_inputs_v2.py`
revalidates the candidate HEAD, clean status, Git tree and tracked-tree digest,
source archive, all three role locks, lock graphs, manifests, every reviewed
vendor blob, toolchain/config, fixture manifest and five probe inputs before
creating a fresh external project.  The generated `api_compat.rs` is a
consumer-only bridge; production format crates are never overlaid or edited.

The preparation receipt binds the complete closure, the 16 ordered routes,
the correctness-only and baseline dispositions, and the owner-confirmed
`signed_batch.selection=immediate_refusal` choice.  Fixture bytes and probe
inputs are staged outside the checkout with canonical, hash-bound receipts.
The generated project uses the legacy `legacy_api` feature branch and the
current Commit-accessor, EditOptions, and XLSB `resolve()` branches in the
same adapter.

`build` revalidates the preparation receipt before invoking the caller's
pinned Cargo shim with offline/locked/jobs=1/incremental=0 and fresh 20-GB
target/TMPDIR boundaries.  `execute` consumes only a canonical v2
capabilities receipt bound to the preparation, build, source closure, adapter,
and fixture receipt.  It requires all 16 routes and typed child evidence when
present, and requires the signed-batch refusal to report zero iterator reads
and zero publication attempts.

The source-only tests cover a positive prepare/receipt/execute pipeline,
candidate identity and role-lock drift, route-set/status drift, immediate
batch refusal, canonical artifacts, and the existing pending/v1 rejection
contract:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_signed_native_policy_v2 \
  tools.test_signed_native_v2_consumer
```

No Cargo, native process, profiler, timing, or metadata command was run for
this record.
