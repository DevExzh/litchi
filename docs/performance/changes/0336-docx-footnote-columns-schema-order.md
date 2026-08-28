# Change 0336: DOCX footnote-columns schema order and namespace correctness

## Scope

This change closes the source-backed DOCX seam for the Word 2012
`w12:footnoteColumns` extension on `w:sectPr`. Insertion of a previously
absent extension follows the schema-valid `sectPr` child order instead of
appending at an arbitrary byte position. Existing standard children, comments,
whitespace, and foreign direct children retain their authored order.

The codec resolves expanded names rather than trusting a caller-selected
prefix. The extension element is required to be a prefixed Word 2012 element,
its required `val` attribute is in the section's WordprocessingML namespace,
and the extension prefix is covered by `mc:Ignorable`. Transitional and strict
WordprocessingML namespace families, locally shadowed prefixes, and namespace
bindings inherited from the owning document part are handled without
serializing inherited declarations into a detached snapshot. Newly generated
qualified names use the exact namespace context of the section.

The package-neutral model keeps absence distinct from an explicit zero. Zero
therefore remains a meaningful request to follow the page's ordinary column
layout, while negative values, malformed decimal values, duplicate values,
non-ignorable extensions, and child content inside the empty extension are
typed format errors.

## Lossless and mutation boundaries

- A value edit changes only the `w:val` value span when the extension already
  exists. Authored prefixes, quote style, whitespace, extension attributes,
  foreign declarations, opaque attributes, and unrelated XML remain
  byte-preserved.
- Inserting or removing the extension rewrites only the required `sectPr`
  seam and namespace/compatibility declarations. It does not regenerate the
  surrounding document part or infer unsupported section mutations.
- A semantic no-op, including a lexical source such as `02` whose parsed value
  is unchanged, returns the original section bytes and an identity patch.
- Commits produce immutable snapshots and source-preconditioned reversible
  patches. Applying a patch to a same-value but different source, malformed
  source, or stale source refuses before publishing a partial result; inverse
  application is atomic.
- Detached snapshots retain inherited namespace and markup-compatibility
  context out of band. The context is bounded and is not treated as authored
  XML, so an untouched snapshot remains exactly the source fragment.
- XML size, depth, element-count, namespace-binding, and retained-context
  limits remain enforced. Ambiguous namespace state, unsafe insertion points,
  unsupported structural content, and malformed XML remain typed refusals;
  there is no lossy fallback.

This is a focused schema-order, QName, and preservation correction. It does
not implement pagination, layout calculation, arbitrary `sectPr` child
reordering, or edits to unknown extension content.

## Validation status

- Crate-scoped rustfmt was applied successfully.
- Independent review found an unresolved Word schema barrier, unsafe ordering
  around unknown direct Word children, inherited `Ignorable` rebinding and
  escaping hazards, and descendant-prefix capture hazards. All four cases were
  hardened with typed refusals or expanded-context handling, with four
  regression tests added.
- Final focused `cargo test -p litchi-docx footnote_columns --lib` passed all
  19 tests.
- Final `cargo test -p litchi-docx --lib` passed all 925 unit tests.
- The earlier full `cargo test -p litchi-docx --lib --tests` run passed 921
  unit tests and every integration target reached before the unrelated
  persistent failure
  `source_backed_file::replacing_the_path_reports_source_changed_without_retargeting`
  at `tests/source_backed_file.rs:270`. The same failure was reproduced in an
  isolated rerun; Change 0336 does not touch that path.
- The initial strict `cargo clippy -p litchi-docx --lib --tests -- -D
  warnings` run found two batch-local `useless_asref` findings; both were
  fixed. That run also exposed unrelated current-tree baseline failures in
  `source_backed.rs`, expectation attributes, and the `sdt_mce_inactive` test.
- Final allowed-baseline `cargo clippy -p litchi-docx --lib -- -D warnings -A
  clippy::redundant_closure_call -A unfulfilled_lint_expectations` passed
  again. The allowances are narrow and cover only those unrelated current-tree
  baseline findings.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-docx --no-deps` passed again.
- The isolated `/dev/shm/litchi-0336-target` tree was deleted after
  validation. Final usage was 36% on `/` with 136 GiB available and 1% on
  `/dev/shm` with 16 GiB available.

The reproduced failure is not attributed to this change. No broader lint or
performance result is claimed from the partial crate-wide test gate.

## Performance claims

`performance_claim: none`

No benchmark, latency, throughput, allocation, RSS, decompression, or process
memory claim is made. Namespace-aware parsing, bounded context capture, and
targeted XML splicing are correctness and preservation properties, not measured
performance results.

## Follow-up

Keep the schema-order oracle tied to the normative WordprocessingML `sectPr`
sequence and retain the typed refusal boundary for unknown direct children
whose safe insertion position cannot be proven.
