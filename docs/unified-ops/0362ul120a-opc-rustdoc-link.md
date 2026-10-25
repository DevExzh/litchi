# 0362ul120a — OPC rustdoc link repair

Status: source-only follow-up to `55a8af6a5a3b30fa139d3a88c112859cc6d042c8`
on `uo2/common-glossary-0362ul120`.  This repair makes no new compiler,
Cargo, rustdoc, runtime, native, or performance claim.

## Retained failure

The suffix glossary packet is retained at
`/home/zhuhe/code/litchi-worktrees/.coord/uo2-coordinator-replay-integration-f5-glossary-suffix-a8d32513b8`.
Its rustdoc attempt was bound to source commit
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b`, exited 101, and produced a
20,177-byte log with SHA-256
`7ed9380e04cd62ef609ee4e7ca9fe113daf7c46d2361259b405adacbdcaca9bb`.
The command used `-D rustdoc::broken-intra-doc-links`; the failure was:

```text
error: unresolved link to `OwnedXmlPart`
    --> crates/litchi-opc/src/package.rs:2509:52
```

This was a documentation-scope error, not a missing type.  `OwnedXmlPart` is
publicly re-exported at `litchi_opc::OwnedXmlPart`, but `package.rs` did not
bring that name into lexical scope for the unqualified intra-doc link.

## Repair and audit

The link at `OpcPackage::try_add_authored_xml_part_bytes` now uses
[`crate::OwnedXmlPart`], preserving a live link to the public type rather
than escaping the reference or weakening the rustdoc lint.  The nearby
publication documentation was audited: the other `OwnedXmlPart` and
`SourceXmlPart` references in this block and the later source-materialization
block already use fully qualified `crate::` paths.  No consumer, API, limit,
allocation, or publication behavior changes.

The suffix packet's source/result anchors remain historical evidence only;
its `rows.json` SHA-256 is
`dd4f59bedeb3556ebc33d210aa4f459ae7e48a8935646fedadbd5bece5c67c9b`, and its
failed `result.json` SHA-256 is
`ddf659bdc19d24a7f79e34164fd9f2d8024fc90640dc1fc4bf575674dfe6c76f`.
The coordinator must rerun the OPC rustdoc and glossary suffix against the
composed source; this record does not relabel the retained failure.

## Source-only checks

`rustfmt --edition 2024 --check` on the changed Rust path and `git diff
--check` pass.  The staged repository-size and binary guards are required at
commit time.  No Cargo, compiler, rustdoc, metadata, runtime, native, or
performance command was run for this repair.
