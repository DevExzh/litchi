# 0362ul89q — PPTX keep-going content-consumer diagnostics

Date: 2026-10-25. This is a source-only repair of the assigned keep-going
diagnostics from the actual F5 attempt at
`15d8d53a0cf6d27f350b866463b4235bf26b08d9`. It does not claim a new Cargo,
runtime, or qualification result.

## Retained diagnostics and source identity

The immutable external capture is
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-15d8d53a0c/`.
Its all-features/all-targets log is 100,739 bytes with SHA-256
`a31129804840cddda84361dc4201852856d55afc25de7f0d49cc7d5868295084`.
The diagnostic index is 18,586 bytes with SHA-256
`c25bee0ed0c486cd54e2d132c928509bfb08484991437136c53c403104b5d3be`.
It retains 130 diagnostics in total; 34 belong to this leaf's assigned
consumer/test paths. Opened tests and other compiler-owner paths remain
outside this repair.

The source repair is commit
`9568aaabf71d60e47e71d123f865762465d07f15`, based on
`15d8d53a0cf6d27f350b866463b4235bf26b08d9`. The external source patch is
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-content-consumers-89q-9568aaabf7.patch`,
14,412 bytes, SHA-256
`5c9d067542d0c78d14a6c8efdae5ac0f1dabba4d2c8b165be8bfd2dea5661e10`.
It contains exactly the ten assigned source paths. No Cargo, build, rustdoc,
or native command was run.

## Repairs and preserved semantics

Expression-valued expected observations in modern-comments, notes, controls,
and OLE test matchers now bind the actual value and compare it in a guard;
the exact observed and maximum values remain asserted. Test-only comments and
notes reexports were narrowed to real consumers, and stale qualified paths
were removed without restoring default-limit wrappers or adding lint
allowances. Notes tests now use the crate result alias, the replay index import
is explicit, and slide-root proof fallback tests pass the selected
`Limits::DEFAULT` profile required by the existing helper.

No production limit policy, codec algorithm, refusal classification, or
selected `*_with_limits` route changed. The edits cover only modern-comments
tests, notes codec/module/package/tests, comments module imports, media-part
tests, and the three assigned embedded-control/OLE codec test modules.
