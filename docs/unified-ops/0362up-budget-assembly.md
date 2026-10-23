# 0362up — Native budget diagnostic assembly

This coordinator branch composes the shared signed-edit, XLSX feature, XLSB,
PPTX and DOCX budget repairs for compiler feedback. It is diagnostic staging,
not acceptance into the integration or trunk branches; review fixes and
combined semantic/performance gates remain required.

Conflict resolution retains immutable OLE Package/PackageState and adds an
inert read-only signed Probe through OpenOptions. The probe keeps the shared
vector source representation, bounded protection policy and no mutation
capability. Ordinary XLS publication keeps the current immutable owner/edit/
commit flow. PPTX retains recipient OOXML policies, XLSB retains required raw
record policies, and DOCX retains signed coalescer selection and authored
operation counts while discarding staged proofs before canonical replacement.

The existing OPC PartMetadata::payload_len supplies admitted deferred lengths
without inflation. Both XLSX and XLSB use that method; duplicate new size
getters were removed during composition.

The first native all-target/all-feature check on `27f185622d` stops on an
incoming PPTX helper referring to the obsolete appended_metadata field. It
now reads the retained policy_appended entry payload lengths.
[native-check-v1.json](results/0362up/native-check-v1.json) records the exact
source, command, external logs and hashes. No native crate pass is inferred
from the stopped build.

The second check on `3ae4cb2420` stops in OPC after 125.55 seconds: the
metadata-only topology planner added a state field without the revision
initializer and omitted a required source policy argument. The subsequent
PPTX proposal replaces that planner with affected-member metadata sizing
and removes both mismatches. Independent review and combined compilation
remain required. [native-check-v2.json](results/0362up/native-check-v2.json)
records the exact command and failure; no dependent-crate pass is inferred.

The third check on `c9bf9c8c40` reaches the native implementations and
reports 21 compiler diagnostics: PPTX wire-length helpers and current metadata
APIs, XLSX retained source projections, DOCX XML views/error conversions and
commit ownership, and XLSB metadata planning. The complete batch is assigned
from that exact assembled source.
[native-check-v3.json](results/0362up/native-check-v3.json) retains the
93.21-second failure. The dependent facade compilation was not reached.

## Scoped review after composition

The diagnostic branch composes the following reviewed repairs without yet
claiming combined compiler/runtime or performance acceptance:

- DOCX `90641c24e8` plus composition fixes `d68a6a05fd` and `2145666675`:
  selected main-XML admission precedes capture, default edits avoid extra
  metadata probes, and ordinary coalesced commits retain native operation
  counts. Signed canonical no-ops retain their authored count.
- XLSB `6521e6b07a`: explicit content-types capture includes both canonical
  digest endpoints and reuses the changed-part census. Grouped cells reuse
  their precomputed native patch charge; current required raw policies and
  admitted deferred payload lengths remain in the composition.
- XLSX features `5ae87e6e08`: manifest charge mirrors actual retained pairs;
  exact no-ops charge the returned source owner twice. Chart staged/expanded
  ledgers, range admission and endpoint accounting passed scoped review.
- PPTX presentation `6cb74b378e`: affected-member lexical sizing preserves
  exact wire framing and skips content-types reads when no override changes.

Review returned the shared XLSX cell/facet/join ledger and metadata capture
ordering for further repair. Compiler adaptation `4cd986a67c` is assembled
for diagnostics; its PPTX metadata path still commits the common native
properties edit before enforcing the selected cap. Exact native metadata
admission is being repaired with the shared XLSX consumer. No public
`finish`/snapshot escape or conservative package-size multiplier is accepted
as a substitute for the retained-payload contract.

Builds are temporarily held while the coordinator runs workspace rustdoc
alone in a separate capped target. Final combined compilation and semantic
tests follow the remaining fixes.

The fourth combined check on `152eb8a53b` stops after 14.74 seconds on
four common metadata preparation errors: a missing no-core branch, boxed
field collection operations and an unused sizing pattern. The narrow compiler
repair preserves the prepared values and does not resolve the outstanding
semantic/performance returns.
[native-check-v4.json](results/0362up/native-check-v4.json) retains the
failure. [0362upa](0362upa-budget-review-return.md) records the required return
work; dependent native/facade and no-feature checks were not reached.

## Stable semantic endpoint accounting

Derived parsed cell and shared-string read caches are not semantic snapshot
payload. Warming those caches must not alter the selected retention budget.
Decoded OPC payload buffers, physical source archives and fixed package data
remain charged, and repeated source/target endpoints remain independent.
The metadata projection therefore uses the same cache-independent endpoint
metric as the primary editor; the temporary warm-cache subtraction bridge is
removed. This is not a general allocator or parser-memory limit.

Metadata proposal `c67826a6bf` passes scoped API, provenance, endpoint,
source-preflight, refusal-order and fidelity review. Rich-transfer successor
`044c37c519` passes scoped review, including shared-index dependencies,
picture-only targets and removal followed by another transfer. Facade compiler
batch `6995c497f5` also passes independent review and its strict compiler run.
All are composed for diagnostic validation. Cell normalization/atomicity,
two PPTX runtime failures, strict combined validation and performance gates
remain required; no full-wave acceptance is inferred.

The tenth combined check on `2cadd934b7` passes strict all-target/all-feature
native/facade checking (238.42 seconds) and the facade without default
features (21.96 seconds). The metadata-focused library run passes ten tests
and fails one cache-warming setup assertion: the shrinking and expanding
fixtures start from an empty workbook rather than the populated helper.
The correction uses that existing helper and retains every exact-boundary
and one-below assertion. [native-check-v10.json](results/0362up/native-check-v10.json)
retains the full terminal results and log hashes. Runtime qualification and
the returned cell normalization work remain required.

The corrected metadata fixtures and reviewed PPTX successor `ae9a31849b`
are composed on `4470ecce16`. All eleven metadata-focused XLSX library
tests pass (73.80 seconds including compilation), and all 27 PPTX facade
tests pass (877.53 seconds including compilation; 1.09 seconds in tests).
This closes both the exact native/coalescer budget failure and empty
`/metadata` no-op failure. No lint suppression or compiler overlay is used.
[native-runtime-v11.json](results/0362up/native-runtime-v11.json) records
the commands, source, log and debug executable hashes. Shared cell
normalization remains returned; these passes do not constitute final
combined runtime, full-gate or performance acceptance.

## DOCX, PPTX and ODF runtime correction

Focused all-feature facade validation passes on `59637bd50c`, including
DOCX module/error cases, malformed inputs, exact signed DOCX paragraph
no-ops and cancellation, and exact signed PPTX shape no-ops and cancellation.
The DOCX correction uses one authoritative bookmark-aware scan and retains
typed structural/limit errors; effective owners keep their physical source
policy. This is a focused runtime pass, not complete composition acceptance.
[docx-pptx-odf-runtime-v14.json](results/0362up/docx-pptx-odf-runtime-v14.json)
retains exact source, commands, test summary and external artifact hashes.

## Complete facade runtime pass

The full all-feature `litchi-ops` library passes all 1,002 tests on
`bf09ddfd9c`, closing the nine failures returned in v12. The run takes
202.67 seconds including compilation and 40.73 seconds in tests. XLSB
then passes 771 tests and returns one transitive replay provenance assertion:
both retained original content-types images differ from effective publication.
The remaining fixture repair must preserve exact published archives and
untouched payload/relationship images. Final native DOCX, the complete gate,
inventory and performance qualification remain required.
[facade-xlsb-runtime-v17.json](results/0362up/facade-xlsb-runtime-v17.json)
retains both commands, test summaries and external log identities.

## Complete DOCX runtime pass and composed lint correction

All 1,519 all-feature DOCX library tests pass on `47caed4627`, including
strict whole-cell refusal and marker-preserving paragraph/replay routes.
Workspace Clippy then reports four OPC lints. Borrowing sink lifetimes now
end in lexical scopes after guard-error extraction, generic constraints are
co-located without changing them, and a redundant usize conversion is removed.
No lint allowance or policy change is introduced.
[docx-clippy-runtime-v18.json](results/0362up/docx-clippy-runtime-v18.json)
retains both complete commands and log identities.

## Complete XLSB runtime pass

All 772 all-feature XLSB library tests pass on `73a38420b8`. The final
transitive replay checks preserve logical member payloads and untouched
source images under the existing OPC writer contract; same-owner and
unreferenced replay retain exact archive comparisons. Workspace Clippy
then reaches seven ODS forwarding lints already corrected on the current
main integration. The coordinator is composing that reviewed source before
further lint validation. [xlsb-clippy-composed-v20.json](results/0362up/xlsb-clippy-composed-v20.json)
retains the complete source and log identities.

## Complete composed lint pass, remaining source corrections

The all-target/all-feature workspace Clippy run with `--keep-going` completes
on `173dd53304` and returns 15 XLSX library/test lints, one DOC redundant
forwarding closure, and the Pages fixture example inclusion errors. The DOC
callback now passes directly without changing invocation order or error
handling. XLSX corrections must retain capture reservation and restoration
semantics; the Pages example must use its fixture helper without invalid
inner documentation or unused test imports. No lint waiver is accepted.
[clippy-composed-v22.json](results/0362up/clippy-composed-v22.json) retains
the complete command and external log identity. A successful final composed
lint run remains required.

## Shared-type composition lint pass

On `ad52437f5e`, the complete all-target/all-feature workspace Clippy run
finishes with three remaining lint sites: automatic dereference in XLSX
facade chart admission and two redundant Diagnostics qualifications in a
PPTX test seam. The qualifications became equivalent after the shared-type
migration. Only their spelling changes; no admission, error, or data flow
changes. [clippy-shared-types-v23.json](results/0362up/clippy-shared-types-v23.json)
retains the complete command and external log identity. The composed glossary
and final-inventory tool suites pass 54 tests; the hardened IWA qualification
suite passes 20 tests.
