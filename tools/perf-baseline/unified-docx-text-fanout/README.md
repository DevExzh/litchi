# Paired DOCX paragraph text fanout

Run from the repository root, with no concurrent build/test/measurement jobs:

```sh
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=/tmp/litchi-unified-0089-target cargo build --release --locked --manifest-path tools/perf-baseline/unified-docx-text-fanout/Cargo.toml
taskset -c 0 /tmp/litchi-unified-0089-target/release/unified-docx-text-fanout-baseline 9 > /tmp/paired-1.json
taskset -c 0 /tmp/litchi-unified-0089-target/release/unified-docx-text-fanout-baseline 9 > /tmp/paired-2.json
python3 tools/perf-baseline/unified-docx-text-fanout/summarize.py /tmp/paired-1.json /tmp/paired-2.json
```

The frozen routes are `docx_eager_leaf_text_fanout_v1`,
`docx_source_leaf_text_fanout_v1`, `docx_eager_unified_text_fanout_v1` and
`docx_source_unified_text_fanout_v1`. Each generated corpus contains 1, 10, 100,
1,000 or 10,000 identical direct body paragraphs, each with one `seed` text run.
The corpus byte sizes and SHA-256 values are enforced by `CORPORA` in main.rs.
Changed corpus or route semantics require new route names.

Both owners start from identical in-memory DOCX bytes. The leaf leg makes one
`replace_body_paragraph_texts` call; the unified leg makes one closure `set`
using a precompiled `/body/p` query and text `updated\ttext\nline` (actual tab
and LF). The native target vector and unified properties are prepared outside
timing. Query execution, target proof, property admission, staging and commit
are inside timing. Opening, query compilation, serialization and destruction
of the retained commit are outside timing. Source-backed owners use memory,
so filesystem revision checks are not represented. This is one fanout operation,
not a command batch with per-item results. The one-target case uses the scalar
unified route and is still compared with the native bulk method.

Untimed verification warms both owners and requires exact serialized output
bytes, native/unified leaf ledger counts equal to the target count, and one
unified operation. Then each invocation runs nine ABBA rounds by default,
18 samples per leg. Adjacent sample pairs in each array belong to the same
round; the native pair surrounds the unified pair. The reported median is the
upper middle sample, consistently used by the Rust and Python runners.

At 10,000 targets the native composite replacement capacity of 4,096 is
exceeded. Both typed refusal paths are checked, and no timings are reported.
The harness does not split the request or treat refusal as successful scaling.

Keep the lockfile. Record the dirty source fingerprint, compiler, machine,
affinity, binary size/SHA-256 and build command. Do not commit the binary.
This baseline does not establish before/after improvement, allocation/RSS,
command batching, real-world documents, file-backed ownership, open/save,
other formats or full performance acceptance.

Record 0091 uses `without-body-cache.patch` as an isolating control. It disables
reuse of the body handle during target admission, preserving every per-node
proof and source fence. It is not a historical source checkout. On matching
source, apply with `git apply --check` then `git apply`, build and copy the
control binary outside the checkout, reverse with `git apply -R`, and build the
candidate. Run control/candidate/candidate/control invocations with nine ABBA
rounds each and no concurrent builds/tests. Run summarize.py separately for the
two control and two candidate files. Preserve the patch, source fingerprint
and both binary identities with the results. Do not leave the control applied.
