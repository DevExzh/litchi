# 0957 bounded provenance packet

This packet records the frozen 0957 source integration, prior-batch provenance,
allocation-design review, and custody corrections before the final allocator
build/capture freeze. It contains no allocation result, comparison result,
adoption claim, binary, source manifest, capture, or final performance report.
The final build is still external and active; root owns the eventual final
source/build/tool/reduction bindings.

The frozen integration is `source-v1` at external
`/home/zhuhe/litchi-performance-0957`, based on commit
`02449f394529766d76794332204c27c4022931fc`. Its external source manifest is
`source-manifest-v2.json`, 2,267,268 bytes, SHA-256
`f5fd1ee75b734bb7efaa1882be0397defea98c76c835be0bdf1f565c573ef617`.
The integration has 10,932 files totaling 273,429,543 source bytes. The six
frozen overlay paths and exact pre-build receipts are recorded in
`source-build-provenance-v1.json` and the copied integration-freeze receipts.

The packet preserves the 0956 progress/root-start record with 35 normative
input hashes, the v2 commit-drop source trace as the primary attribution note,
the allocation design and final source review, and the post-build draft-write
history. The post-build draft and benchmark drift was retained as historical
unapproved evidence and then restored to the pre-build source manifest. The
allocator library target label correction is metadata-only and preserves the
original receipt.

The active build/tools are intentionally not copied or bound in this
pre-freeze packet. The provenance descriptor records exact source reconstruction
and build commands. Source reconstruction refuses an existing destination:
create a fresh `$REBUILD` with `test ! -e "$REBUILD" && mkdir "$REBUILD"`;
never remove an existing tree as part of regeneration. Large manifests,
binaries, targets, logs, and any eventual capture remain under the external
0957 packet.
