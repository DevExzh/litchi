# Source and measurement review

The root inspected current HEAD 25b5516b0 and the complete staging/commit data
path before editing. The previous turn completed and committed 0440. The
accepted ADR tree is unchanged from the prior complete read; no new exception
is required for sharing an immutable private source projection.

The only production changes are the source comparison field and constructor
argument in authoring/mutable.rs and the snapshot's private constructor call
in authoring/edit.rs. The editable Vec still deep-clones the validated slides.
The source Arc is never mutated; retained_page still uses the same origin and
semantic comparison before returning exact original markup. The transaction
already owns a source snapshot, so sharing does not extend the source lifetime.
Empty fallback behavior and incomplete source-page coverage refusal remain.

The existing isolation test now checks pointer sharing of the immutable
projection, distinct editable storage, an unaffected peer draft, and retention
after the input handle and modified draft are dropped. The coverage-refusal
test uses an incomplete Arc projection. All 352 ODP tests and strict owner
Clippy pass. Public interfaces, security checks, source limits, namespace
semantics, preservation, publication and readback are unchanged.

Two setup failures remain retained. The first A1 invocation omitted required
repo/custody paths and exited in argparse before any workload ran; the corrected
invocation captured the only A1 measurement set. The first candidate compile
exposed the crate's unnecessary-qualification lint after an Arc import and a
Vec-versus-Arc assertion mismatch. The root removed that import, qualified the
new production uses, and compared slices in the test before rerunning. No
measurement uses the failed candidate source. An initial derivation without
--write validated the matrix but did not create summary.json; the subsequent
explicit --write invocation retained the result. No raw measurement was changed.

All main reports and four profiles were captured serially against their exact
retained binaries and full source manifests. Source switches happened only
after the prior CPU handle returned terminal status. No measurement attempt is
excluded or substituted. Profiles retain raw data and both text conversions,
including all symbolization warnings.

The original practical gate failed. The separate acceptance-review.md explains
the post-hoc memory decision and the root's omission of a peak-memory criterion
from the original gate. The 720 samples, failed gate, slight normal latency
cost and baseline tiny p99 repeat flag remain visible. This is a scoped memory
improvement, not a normal latency, RSS or bounded append result.

Final validation passes 368 harness tests with one ignored, rustdoc with
warnings denied, explicit-file pinned formatting and complete crate boundaries.
The current full source manifest equals the measured candidate, both retained
source copies match their roles, and the GOAL hash, accepted ADR tree and sealed
0440 evidence are unchanged. The copied oracle is byte-identical to 0440.

Portable copied controls pass before and after cleanup. All 11 precleanup
and 12 postcleanup corruptions are rejected after refreshing inventories,
including a false claim that the original practical gate passed. Cleanup
removes the explicitly owned binary directory (1,830,211,696 regular-file
bytes), preserving both shared build-cache directory identities and the GOAL
digest. Generated check.py bytecode was also removed before sealing. The final
bundle verifies after the binaries are gone, including reproducible summaries
and measurement tables from the compressed artifacts.
