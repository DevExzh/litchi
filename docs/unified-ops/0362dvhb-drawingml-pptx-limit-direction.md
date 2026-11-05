# 0362dvhb — PPTX subordinate-reader limit direction

This return packet follows `0362dvha` and corrects the PPTX shape facade's
subordinate-reader policy. `Common::transform` and `Common::text_body` both
admit their source fragments under the retained scene `input_bytes` budget;
the scene's `output_bytes` budget remains the owner-level processed-output
bound enforced while the scene is indexed.

The mapping is covered by asymmetric input/output limit regressions in the
shape model so a future change cannot silently route a source reader through
the output budget. No compatibility API or alias was added.

Validation receipt for source `ca738e3689` (the import-only follow-up to
`00fb949e54`):

```text
source (deleted-coord-artifact)
export CARGO_TARGET_DIR=(deleted-coord-artifact)
export CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 CARGO_NET_OFFLINE=true
cargo test --offline -p litchi-pptx shape::model::tests::subordinate_readers_use_the_scene_input_budget --lib
cargo test --offline -p litchi-pptx --test shape_text_body owned_lazy_body_keeps_the_caller_limits
cargo check --offline --all-targets -p litchi-pptx
```

All three commands passed under the 20 GiB memory cap with one build job and
offline resolution. These commands did **not** pass `--locked`, so this is
not a pinned-lock receipt. The ignored leaf `Cargo.lock` observed after the
validation sequence was 139,592 bytes with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`;
the pre-run lock hash and per-command post hashes were not captured.

No external raw command log was retained; the output was observed in the
interactive validation session. The shared external target was
`(deleted-coord-artifact)`;
its current post-inspection apparent size is 777,184,046 bytes. Command-time
pre/post target measurements and log byte/hash receipts were not captured.
This is diagnostic leaf validation only (`6d`); assembled `4100` workspace
validation remains coordinator-owned and required. No native workload or
ABBA run was performed.
