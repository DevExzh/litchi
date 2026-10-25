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
inputs are staged outside the checkout with canonical, hash-bound receipts;
the staged manifest is also required to equal the immutable source fixture
identity, so changing only the staged copy cannot change the input contract.
The historical v1 probe is recorded as a source input only.  Preparation
deterministically transforms that pinned source into a v2 producer: it
changes the schema to `litchi.signed-native.probe.v2`, adds the route
dispatch and the compiled `batch_refusal` producer, and retains the API
compatibility exercise.  The generated source is re-created and byte-checked
on every revalidation; no v1 receipt is wrapped as v2.
The generated project uses the legacy `legacy_api` feature branch and the
current Commit-accessor, EditOptions, and XLSB `resolve()` branches in the
same adapter.

`build --api-branch current_api|legacy_api` revalidates the preparation receipt
before invoking the caller's pinned Cargo shim with
offline/locked/jobs=1/incremental=0 and fresh 20-GB target/TMPDIR boundaries;
the selected branch is recorded in the build receipt, so an uncompiled cfg
branch is never counted as coverage.  The v2 runner then invokes the actual
binary once for every route and once for `batch_refusal`.  Each child receipt
binds the exact argv, binary, fixture root/manifest/hash, working directory,
owned process group, bounded stdout/stderr artifacts, timeout and exit status;
the parsed child JSON is compared exactly with the producer's stdout value.
The
runner's batch receipt similarly binds the producer's raw source-before and
source-after bytes.  `execute` consumes only these canonical v2 capabilities,
bound to the preparation, build, source closure, adapter, runner and fixture
receipt.  Each route must carry one typed result for every staged fixture,
bound to its manifest name, format, origin, and producer source digest where
available.  It requires all 16 routes with route-specific semantic evidence,
and requires a signed-batch `DocumentProtected` proof with `item=null`,
unchanged before/after source artifacts, zero iterator reads, and zero
publication attempts.

The source-only tests cover a positive prepare/receipt/execute pipeline,
candidate identity and role-lock drift, route-set/status drift, immediate
batch refusal, canonical artifacts, a process-shaped v2 runner that executes
all routes and the batch producer, and the existing pending/v1 rejection
contract:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_signed_native_policy_v2 \
  tools.test_signed_native_v2_consumer
```

This command runs 22 source-only tests.  The runner test uses an executable
Python process as a deterministic source-level stand-in for the future built
binary; it exercises real subprocess capture, process ownership, raw stdout
binding, and all route/batch dispatch without claiming native correctness.
No Cargo, native binary, profiler, timing, or metadata command was run for
this record.  The checked-in candidate binding remains pending, so a real
run still requires the coordinator's separately reviewed source freeze and
guarded build/smoke packet.
