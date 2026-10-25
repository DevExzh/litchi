# 0362ul72a — DrawingML v1 frozen-source boundary

Date: 2026-10-25. This source-only follow-up keeps the version-one
DrawingML receipt contract honest while version two becomes the active current
source protocol. It does not run Cargo, rustdoc, a native probe, or a timing
workload.

## v1 is bound to its approved historical input

The v1 consumer adapter remains the exact six-spelling policy recorded by
0362ul57b and earlier records. Its source-bound tests no longer read whatever
`tools/perf-baseline/src/lib.rs` happens to be at the current repository tip.
They read the approved standalone source blob directly from fixed ref
`805c08facf9ff7c92c7f05b66bde1e2bdb8524f2`:

```text
path:       tools/perf-baseline/src/lib.rs
Git blob:   2ec7f936b22bef7468983134b0b5f05386f7f8a8
SHA-256:    4acd072fd68c094e24524862194da9676260e1fb25f6ea3dbb539197d2838b5b
bytes:      3,088,209
```

The tests verify the fixed ref, blob id, byte count, and SHA-256 before
applying the historical six substitutions. This preserves the meaning of old
v1 receipts and prevents an unrelated current `lib.rs` from silently becoming
their input.

The v1 runtime still rejects a candidate whose consumer adapter is not that
frozen six-spelling source. In particular, the later RTF lane source at
`dc2b1238a4654cccd1ae86d64215b88aabb8b740b25c54266123247d92799e51` is not
accepted by v1 merely because its other routes are compatible. No historical
receipt is relabeled or refreshed to that source.

## v2 remains the current-source path

The v2 preparer from 0362ul72 does not reverse consumer spellings. It copies
the candidate's current probe/routes and materializes the historical
`lib.rs` from the fixed `472b...` blob. Its candidate binding remains
intentionally incomplete, so execution fails closed until the final source
commit, manifests, and complete tracked vendor-tree digest are committed into
the policy. The v1 frozen-source tests and v2 pending-policy tests therefore
both pass against a later current tree without fabricating a future candidate.

Focused checks:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_drawingml_raw_qualification \
  tools.test_drawingml_raw_policy_v2
```

The combined source-only suite passes 24 tests on this leaf. No performance or
native qualification claim is made.
