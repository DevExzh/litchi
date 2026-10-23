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
