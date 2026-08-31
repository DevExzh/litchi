# Change 0338: XLSX source-backed shared-formula group transaction

## Scope

This change closes the source-backed XLSX seam for editing an authored shared
formula group. The public operation is explicitly master-addressed: the
caller supplies the address of the group's master cell and a replacement
formula, and the operation refuses when the supplied address is a follower.
The master address is not inferred from a follower, a calculation-chain
entry, or a caller-selected shared-formula index. The authored shared-formula
identity, master `ref`, and all participating cells must be resolved before
any output is constructed.

The group is accepted only when its membership is a complete rectangular
range described by the master `ref`. Every cell in that rectangle must be
owned by the same worksheet, have the same shared-formula index, and have a
valid authored shared-formula role. The master retains the shared formula's
authored topology; followers continue to be shared-formula followers rather
than being expanded into independent formulas. Formula text is replaced at
the master, and no cached result is retained for the changed group because a
cache from the previous formula is no longer authoritative.

The aggregate group bound is 256 members, including the master. This is a
transaction bound, not a promise to process an arbitrarily large worksheet.
The rectangular-membership proof, formula validation, and all retained XML
and relationship state are completed within the configured read and output
limits. A group that is larger than the bound, sparse, overlapping, split
across worksheets, duplicated, or otherwise ambiguous refuses before
publication.

## Lossless and mutation boundaries

- A semantic no-op, including a master-addressed replacement whose typed
  formula and authored shared topology are already unchanged, returns the
  original package bytes and an identity patch. It does not rewrite any
  worksheet cell, alter `calcPr`, remove `calcChain.xml`, or change package
  topology.
- An effective group edit rewrites only the proven shared-formula group. The
  master remains the master, every follower remains a follower, the complete
  rectangle remains represented, and unrelated cells, authored dimensions,
  styles, strings, XML whitespace, relationships, and raw parts remain
  preserved.
- The old cached values for the changed group are removed. The resulting
  cacheless shared formula is not evaluated by Litchi and does not fabricate a
  replacement cached result. Any stale `<v>` payload associated with the
  edited group is removed as part of the same worksheet rewrite.
- Every effective group edit invalidates workbook calculation properties
  (`calcPr`) and removes the workbook calculation-chain relationship and its
  `calcChain.xml` part, including the content-type entry, as one atomic
  source-topology transaction. The calculation chain is not partially
  regenerated from the edited group.
- Patches are immutable, source-preconditioned, and reversible. Forward
  application restores the same shared master/follower topology with the new
  formula; inverse application restores the prior formula, caches, and
  calculation-chain topology through the same validated operations. A
  changed or semantically incompatible source refuses before a partial
  package or partial patch is published. Changed-source inverse restoration
  is semantic and does not promise authored manifest lexical order.
- A master-addressed request against a follower, a missing master, a formula
  group without a provable `ref`, or a group whose complete membership cannot
  be shown is a typed refusal. The operation never silently expands followers
  into scalar formulas or edits only the selected cell.
- Strict group proof is applied only to a selected shared-formula edit.
  Oversized or otherwise unsupported shared groups remain guarded against
  ordinary member edits but do not prevent an unrelated scalar edit elsewhere
  in the worksheet from preserving their source bytes.
- Signed packages refuse mutation, including calculation metadata
  invalidation and calculation-chain removal. Source freshness is checked
  before output; cancellation is observed at bounded source capture,
  rectangular-membership validation, XML rewrite, and topology publication
  boundaries.
- Read, retained-byte, XML, relationship, part-operation, group-member, and
  output limits apply to the aggregate transaction. The 256-member aggregate
  bound is enforced before mutation. Any limit breach refuses atomically and
  leaves no partial package, patch, worksheet, or topology update.

This is a source-backed shared-formula transaction for an existing, scalar
shared-formula group. It does not evaluate formulas, repair malformed
workbooks, normalize unrelated XML, create cells, resize the group, change
the authored range, or claim support for groups whose dependency closure is
not bounded and proven.

## Validation status

- `cargo fmt --package litchi-xlsx`: passed.
- The complete `source_backed_cell_values` integration target passed: 48
  tests, zero failures, one test thread.
- The exact test command was:

  ```sh
  CARGO_TARGET_DIR=/home/zhuhe/CodeProjects/.cargo-targets/litchi-0338-target \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_TEST_DEBUG=0 \
  cargo test -p litchi-xlsx --test source_backed_cell_values -- \
  --test-threads=1
  ```

- Read-only static review identified and closed two preservation findings:
  strict proof no longer rejects an unrelated edit because of an untouched
  oversized group, and namespace declarations are no longer misclassified as
  unsupported formula attributes. Both paths have focused regression coverage.
- Broader XLSX library/integration, strict Clippy, and rustdoc gates were not
  run. No repository-wide validation claim is made after the prior host OOM.
- Scoped `git diff --check` passed for the production and regression files.
- The isolated target was deleted after testing. At finalization the root disk
  had 136 GiB free and `/dev/shm` used 53 MiB; `/tmp` remained an unrelated
  16 GiB tmpfs and was not used or modified by this change.

All Cargo work was serialized on one disk-backed target. No parallel rebuild
or tmpfs Rust target was used.

## Performance claims

`performance_claim: none`

No benchmark, latency, throughput, allocation, RSS, decompression, or process
memory claim is made. Bounded rectangular membership, exact no-op
publication, and targeted worksheet/topology edits are correctness and
preservation properties, not measured performance results.

## Follow-up

Array formulas, data-table formulas, external-workbook formulas, and any
formula dependency closure that cannot be proven within the owning package
remain deferred. Future support must preserve each authored group kind and
range, retain the master-addressed API boundary, enforce bounded complete
membership, invalidate calculation metadata transactionally, and keep the
same source-preconditioned inverse and cancellation guarantees.
