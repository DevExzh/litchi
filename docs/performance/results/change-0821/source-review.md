# 0821 source review — ordinary OOXML save and durability boundary

## Reviewed source and quality base

The review is against the repaired current base `8312aaa29b` and is limited
to source and protocol boundaries. No Rust source or runtime harness source is
changed by this packet. The 0820 quality repair changed only the test portion
of `tools/perf-baseline/src/bin/support/counting_allocator.rs`; it corrected a
process-wide live-byte assertion and did not change the allocator
implementation, ordinary-save harness, or production crates.

The directly relevant current source hashes are:

| File | SHA-256 |
| --- | --- |
| `tools/perf-baseline/src/ordinary_save.rs` | `09d9a2308e2545d038ef785a0770b3e21d26e7379f356bdd6671de65cebddd12` |
| `crates/litchi-core/src/durability.rs` | `229626c04d11dc11e09c0535fa99c228ee39a1b30d070d688560ca35ac28433f` |
| `crates/litchi-opc/src/atomic.rs` | `75a85ce671c340a4ad1dbceb766441d0c4936beb5412db8e022bdbdeb16dbc4f` |
| `crates/litchi-docx/src/package/codec.rs` | `c7d37ad55763021ba69d1ff47b8bfe2be077c54f20dbaa155ac7415674aff106` |
| `crates/litchi-xlsx/src/package.rs` | `10d94bdf99839f5ede8089c04493b55a000b4fa5707bcbe354b0ef96b89bf309` |
| `crates/litchi-pptx/src/package/codec.rs` | `9a0c78fa7c2d77a9e19c442fe75559237fa8334dea46d6668050b55ed7dbb5b0` |
| `crates/litchi-opc/src/pkgwriter.rs` | `eea4f1bef388d19a68d2dd8fd66fffec1879dfa83f2e6104c88429e118cbdf7d` |
| `docs/GOAL.md` | `bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1` |
| `docs/adr/0003-snapshots-edits-and-patches.md` | `9200d3546d91f1604d44a9ade7ad0a6bdea50b6442d12a0517b09db5e8e802ef` |
| `docs/adr/0005-io-memory-and-performance.md` | `8279fd2dc0b074aa104bf0c83df6187d12e18756cd8c8d7f25750b2a997cb0aa` |
| `docs/adr/0006-validation-security-and-compatibility.md` | `ae21189eda0acc9524a7c5a57e87bb3566405ede5c015c87105e01337666f44b` |
| `docs/adr/0011-ooxml-physical-package-ownership.md` | `3a1644536af0b66fed47c1a154e829c0e0008cccb89dec679a0328036d305e04` |
| `docs/adr/0024-current-topology.md` | `cc134b2a4eb8797a3769cf815f1f09618fd26e215bd7305de1c9890093dbb1b6` |

The 0821 packet must retain these values in its frozen source descriptor and
recheck them before capture. The offline readers are a post-capture handoff,
so their final source/analysis descriptors are checked only after all capture
children terminate.

The exact 0820 repair quality evidence is the authority for reused checks:
`repair/quality.json` is `status: pass` with six gates,
`repair/test-summary.json` records 641 passed, zero failed, one ignored across
28 suites, and `repair/source.json` records 9,197 production files. The 35
normative input hashes in `architecture-inputs.json` all match the current
tree at review time. The 0821 packet must carry these as verified references
and retain the distinction between reused quality and fresh release builds.

## Production route

`litchi_core::Durability` is a caller-selected, per-call enum. `Full` is the
default and requests synchronization of the sibling temporary file followed
by the destination parent directory where supported. `FileOnly` retains the
temporary-file synchronization and skips the parent-directory step. `NoSync`
skips both synchronization calls. The level is neither persisted in a package
nor read from environment, global state, a clock, or a hidden executor.

`litchi_opc::atomic::replace_with_durability` performs the common publication
route: destination validation, sibling temporary creation in the destination
directory, complete writer output, sink finish, existing regular-file
permission preservation where promised, optional temporary-file sync, one
same-directory rename, and optional parent-directory sync. A failure before
the rename leaves the destination untouched and removes the temporary file.
Only Full can report a post-replacement parent-directory synchronization
failure. The weaker policies skip their named operations and therefore cannot
report errors from operations they did not attempt.

The OOXML owners forward the same caller-selected level to this physical
owner. `Package::save` in DOCX, `Workbook::save` in XLSX, and `Package::save`
in PPTX call their `save_with_durability(..., Full)` forms. The explicit
forms use the same writer and atomic replacement path. Thus default/full is a
route control; a difference between them is dispatch or run-to-run variation,
not a synchronization optimization.

## Harness boundary

`ordinary_save.rs` uses the same format-specific open, semantic edit, and
save route for every filesystem policy. `run_case_with_durability` accepts a
durability level only for `lifecycle` and `atomic_publish`; edit and sequential
counting phases cannot accidentally claim a durability policy. The lifecycle
clock includes open, edit, and save. The atomic clock includes only
`save_with_durability`/`save`, with open and edit outside it. Readback,
digests, destruction, and destination removal are outside both clocks.

The exporter separately invokes default, explicit Full, FileOnly, and NoSync,
then the sequential sink control. It checks the output digest, reference
identity, source identity, reopen, and edit outcome before artifact admission.
The independent Python artifact audit and ZIP preservation reader must then
check semantic target closure, relationship/content-type closure, untouched
XML/binary members, and output policy identity. These checks are necessary
because equal output hashes alone do not establish lossless preservation.

The report's `timing_scope` is a generic phase description. The authoritative
policy evidence is `save_durability` plus `atomic_publication_steps`. In
particular, an omitted `save_durability` means documented `save` and therefore
Full; it does not mean that no durability policy was selected. Readers must
not infer a synchronization syscall from the generic timing string.

## Ownership and ADR checks

The harness calls public format APIs and does not expose archive implementation
types, physical IDs, locks, executors, or runtime handles. The experiment adds
no filesystem or network provider to production, no ambient thread pool, no
parallel save behavior, and no source or preservation shortcut. It keeps
atomic replacement and sequential sink semantics intact. The independent
artifact gates retain the content-type and relationship closure checks needed
by ADR 0003, ADR 0005, ADR 0006, and ADR 0011.

The 0821 protocol uses one fresh destination per sample. It therefore does
not exercise replacement over an existing regular file or measure permission
copying in the timed path. Warm filesystem/provider caches and one shared host
remain the declared environment. Procfs and allocator observer fields are
diagnostics with fixed controls; they are not operation-local or native timing
evidence.

## Review decision

The production route has no source blocker. The 0821 matrix is valid as a
durability-policy attribution experiment when its four policy fields are
reported exactly and Full remains the documented default. Execution may begin
after the 0821 plan, source descriptor, quality-reuse receipt, and root-owned
drivers have been frozen with current identities. Fresh release-build,
artifact, qualification, and capture receipts are then produced in sequence.
The offline readers are checked after the children terminate and before
analysis/sealing. Any stale 0820 schema or revision in an 0821 driver or
reader must be fixed before its relevant handoff; the explicit 0820
quality-reuse references are expected and remain valid.
