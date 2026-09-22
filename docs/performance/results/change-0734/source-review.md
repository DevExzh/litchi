# 0734 PPT owned-stream source review

This independent review covers the 0734 candidate for the private embedded PPT
editor finish path. It compares the archived baseline with the candidate copy
and the formatted candidate currently applied in the worktree. The review makes
no Rust, Cargo, probe, or measurement changes and runs no build, native, or
profiler command.

Disposition: **approved for the frozen semantic and performance gates**. I
found no correctness blocker in the ownership handoff. The candidate keeps the
existing validation and output boundaries in place, and its duplicate-path
fallback preserves the reachable borrowed behavior. The quality and native
evidence still decide whether the production hunk is retained.

The reviewed source identities are:

| item | path | SHA-256 |
| --- | --- | --- |
| archived baseline | [`source-archive/before/.../finish.rs`](source-archive/before/crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs) | `0de96a64f79de8c5aacda28fb2387b45966ca5736d1c1b1963b16e29826701e1` |
| candidate copy | [`candidate/.../finish.rs`](candidate/crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs) | `b7ded351a1ea0afb48204f90e675e7bba77f8fd14543e3d9049daa4125142931` |
| formatted live source | [`finish.rs`](../../../../crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs) | `31132455baead9b66dbea9a7760cdaac202e45c87e81ce969b6d290a3d158834` |
| candidate notes | [`implementation-notes.md`](candidate/implementation-notes.md) | `77657aae4ab987f3f38b4e2f87865c17cec249da41644b2c7454e61dc0e3eaa5` |

The live source differs from the candidate copy only by `rustfmt` layout. The
0734 constraint snapshot is [`constraints.json`](constraints.json), SHA-256
`13fbdce87065a406b83b58f85ca2f32668a49f4b8306f5012afd2a3092a20863`, and the
source review is anchored to baseline commit `769f9824ce92274643db9f5566c69ff37f094754`.

## Applicable constraints

I read the accepted ADR index and the relevant accepted records: [ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR 0002](../../../adr/0002-crate-topology.md), [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md), [ADR 0005](../../../adr/0005-io-memory-and-performance.md), [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md), [ADR 0008](../../../adr/0008-migration-and-verification.md), [ADR 0010](../../../adr/0010-facade-archive-ownership.md), [ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md), and [ADR 0024](../../../adr/0024-current-topology.md). The 0734 constraint manifest freezes all accepted ADR hashes.

The change stays inside the existing `litchi-ppt` editor and calls the existing
`litchi-cfb::OleWriter` API. It adds no public API, dependency, unsafe code,
cache, executor, ambient I/O, or new persistence policy. The CFB ownership
direction remains downward and the iWork exclusion is unaffected.

## Ownership and ordering

The baseline's `write_package` copied every selected payload through
`OleWriter::create_stream`. The candidate adopts the source layout first, then
counts exact selected paths and moves the consumed `Editor::streams` vectors,
the assembled Document vector, and the updated Current User vector into
`create_stream_owned`. The stream loop retains its original order and the
Document/Current User substitutions retain their original complete-path
matching. `validate_rewrite` still runs on the bytes actually emitted by the
writer.

The order matters for both preservation and errors. Source-layout adoption and
its fallible parse happen before any editor payload is moved. All projected
Document sizes, persist mappings, append construction, UserEdit construction,
and the Current User offset update are unchanged and still precede the writer.
The bounded sink remains the writer's destination, and its limit error mapping
is unchanged. The post-write `validate_rewrite` fence still follows emission
and validates the bytes actually produced. `Reuse` remains the default and
`Rewrite` remains an explicit policy choice; no plan validation, readback,
flush, or final reopen was removed.

`Editor::open` selects the first stream whose leaf name matches each required
PPT stream, while `Editor::streams` retains complete paths. The candidate does
not infer global leaf-name uniqueness. If a selected complete path occurs once,
its replacement vector is moved exactly once. If it occurs more than once, all
occurrences use the old borrowed `create_stream` path, so repeated replacement
and writer overwrite behavior remain the same. Missing selected paths retain
the old behavior: no stream is injected. Other complete paths are moved one at
a time and preserve their insertion order. The private contradiction guards
cannot be reached after the occurrence counts and loop are fixed, but prevent a
future ownership edit from silently panicking.

The move consumes only the editor passed to `finish`, which is already a
consuming operation. The immutable source `Arc<[u8]>` is never modified, and
the fields used after package emission (`document_path`, `current_user_path`,
`collection`, and the limit) remain available for final validation. A writer
failure drops owned buffers through ordinary Rust ownership and returns the
same typed OLE error channel; it does not publish a partial editor state.

## Semantic and resource equivalence

The unchanged finish body preserves the append-only PPT transaction contract:

1. removed mappings are cleared;
2. staged records and an optional rewritten object list are projected under the
   existing `u32` and output-size checks;
3. the incremental Document stream, persist directory, and UserEdit record are
   assembled in the existing order;
4. Current User points at the new edit;
5. the CFB writer emits through the selected layout policy and bounded sink;
6. the final package is reopened and its live mapping is checked.

The normal no-op branch is byte-preserving and unchanged. The candidate's
allocation difference is deliberate: a normal unique stream no longer creates
the second payload vector that `create_stream` used to reserve and fill. This
removes the old artificial `stream payload` allocation failure opportunity. The
remaining path/table reservations and all format, mapping, output-limit,
writer-I/O, and reopen failures stay typed. Under an extreme simultaneous
allocation failure, the low-level resource label can therefore be `stream path`
or `stream table` where the old copied path would have failed first at
`stream payload`; that is the expected failure-surface consequence of removing
the payload allocation, and is not a semantic refusal or a guessed edit.

The source review does not infer native latency, allocation, RSS, or instruction
savings from the source diff. Those require the frozen 0734 baseline/candidate
process measurements and allocation lane.

## Test coverage and residual review notes

The candidate adds focused tests that compare the old borrowed writer with the
owned writer under both `Reuse` and `Rewrite`, compare a changed persisted
record against the complete borrowed finish path, reopen and read back the
replacement, check exact no-op and changed source immutability, retain a typed
projected-stream limit error, and exercise duplicate selected complete paths.
The existing editor mapping, limit, public commit, source-preservation, and
semantic tests remain outside this source file, while the frozen 0734 probe
provides the primary and qualified secondary public-route controls.

Two small test gaps remain advisory rather than blocking: the focused module
does not inject a low-level writer I/O failure, and it does not construct a
duplicate same-leaf stream at a different complete path. The first is covered
by the unchanged writer error boundary and must remain in the broader quality
gates; the second does not enter either replacement branch because matching is
by complete path, while the exact-path duplicate fallback is directly tested.
Neither gap justifies changing production behavior or adding a duplicate
rejection.

No independent review finding requires a source correction. Retain the
candidate only if the frozen output, semantic, negative-control, allocation,
latency, peak-live-byte, and statistical gates in the 0734 packet pass; if the
candidate is rejected on those gates, the source archive remains sufficient to
restore the one owned production file.
