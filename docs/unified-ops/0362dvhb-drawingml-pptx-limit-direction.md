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
source /home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/env.sh
export CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/drawingml-fragments-raw-target
export CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 CARGO_NET_OFFLINE=true
cargo test --offline -p litchi-pptx shape::model::tests::subordinate_readers_use_the_scene_input_budget --lib
cargo test --offline -p litchi-pptx --test shape_text_body owned_lazy_body_keeps_the_caller_limits
cargo check --offline --all-targets -p litchi-pptx
```

All three commands passed under the 20 GiB memory cap with one build job and
the locked offline dependency set. No native workload or ABBA run was
performed.
