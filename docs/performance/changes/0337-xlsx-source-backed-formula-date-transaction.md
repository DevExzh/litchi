# Change 0337: XLSX source-backed formula and date transaction

## Scope

This change closes the source-backed XLSX cell-edit seam for existing cells.
The transaction accepts a typed `Date` value and replacement of a scalar
formula. A scalar formula is written without a cached value; retaining the
old cached result would make the published workbook claim a result that no
longer belongs to the formula. Formula text is kept in the package's ordinary
formula representation and is not evaluated by Litchi.

Every effective cell edit is a workbook transaction rather than a
worksheet-only splice. Formulas elsewhere in the workbook may depend on a
changed value or date, so the transaction invalidates the workbook calculation
properties (`calcPr`), removes the calculation-chain relationship from the
workbook, removes the corresponding `calcChain.xml` part, and removes its
content-type Override. The topology change is made atomically with the
worksheet replacement.

The edit surface remains limited to existing cells with a proven scalar
closure. Shared, array, and data-table formula groups are not rewritten by a
single-cell operation. Group edits require a separate operation that can
prove and rewrite the complete group; until that capability exists, these
requests refuse before output.

## Lossless and mutation boundaries

- A semantic no-op, including a formula or date request whose typed value is
  already present, returns the original package bytes and an identity patch.
  It does not invalidate `calcPr`, remove a calculation chain, or rewrite a
  worksheet.
- A changed scalar formula drops the stale cached value. Every changed cell
  performs the worksheet, workbook, relationship, content-type, and
  physical-part changes as one bounded publication. Untouched raw parts,
  worksheet cells, authored dimensions, relationships unrelated to the
  calculation chain, and package ordering remain preserved.
- The inverse is semantic and source-preconditioned. It restores the prior
  cell contents and prior calculation-chain state through the same validated
  topology operations; it does not depend on manifest member order or a
  particular lexical spelling of `[Content_Types].xml`. Inverse application
  is atomic and refuses when the source has changed.
- A formula edit that needs a shared string table, external workbook, or any
  other dependency outside the selected bounded closure refuses rather than
  silently approximating a cached result or materializing unrelated content.
  A cacheless scalar formula remains valid when no such dependency is needed.
- Unknown cell/formula kinds, unsupported markup-compatibility (`mc:`)
  content, ambiguous extension content, and formula groups whose complete
  membership or authored structure cannot be proven are typed refusals.
  Existing unsupported content remains readable and preservable.
- Signed packages refuse mutations, including a formula-triggered calculation
  chain removal or a date edit. Freshness/precondition checks run before any
  output is published, and cancellation is observed at bounded read,
  rewrite, and topology-operation boundaries.
- Read, retained-byte, XML, relationship, part-operation, and output limits
  apply to the aggregate transaction. A request that would exceed any limit
  refuses without a partial package or partial patch.

This is a source-backed transaction for existing cells. It does not evaluate
formulas, repair malformed workbooks, normalize unrelated XML, rewrite
formula groups, or claim general support for external-link formulas.

## Validation status

- `CARGO_BUILD_JOBS=1 cargo fmt --package litchi-xlsx`: passed.
- `CARGO_TARGET_DIR=/dev/shm/litchi-0337-target CARGO_BUILD_JOBS=1 cargo check
  -p litchi-xlsx --lib`: passed.
- `CARGO_TARGET_DIR=/dev/shm/litchi-0337-target CARGO_BUILD_JOBS=1 cargo test
  -p litchi-xlsx --test source_backed_cell_values`: passed, 39 tests.
- The focused suite covers direct dates, scalar-formula cache removal,
  `calcPr` invalidation, calculation-chain topology removal/restoration,
  semantic no-op byte identity, grouped-formula refusal, stale-source
  refusal, managed limits, signed-source refusal, atomic multi-sheet
  publication, untouched-member preservation, and patch inverse behavior.
- `git diff --check`: passed.
- Static review found that owning patch application could remove the
  calculation-chain Part while another package or Part relationship still
  targeted it. The patch now refuses any additional inbound relationship,
  including ASCII-case-equivalent Part names, before mutating its candidate.
  A focused atomic-refusal regression was added and source-reviewed after the
  OOM cleanup; it was formatted but not executed.
- The broad XLSX `--lib --tests` gate did not complete before the host OOMed
  and was deliberately not restarted. Clippy and rustdoc were also omitted to
  avoid another high-memory crate-wide build. This record does not claim those
  gates passed.
- The 8.7 GiB isolated target was removed after validation. `/dev/shm` fell
  from 8.8 GiB used to 53 MiB, leaving 16 GiB free; the root filesystem had
  135 GiB free.

## Performance claims

`performance_claim: none`

No benchmark, latency, throughput, allocation, RSS, decompression, or process
memory claim is made. Bounded dependency capture, exact no-op publication, and
targeted XML/topology edits are correctness and preservation properties, not
measured performance results.

## Follow-up

Keep formula-group closure and external-dependency support as explicit
capabilities. Any future support must preserve authored group semantics,
invalidate calculation metadata transactionally, and retain the same
source-preconditioned inverse and bounded-failure guarantees.
