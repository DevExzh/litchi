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
