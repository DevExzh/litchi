# Log sections for change 0662

Four paragraphs for the coordinator to merge into the shared performance logs.
This packet deliberately does not edit the shared rollups.

## For `HOTSPOTS.md`

## 0662 — changed members now deflate under a bounded write wave; the read-session rows remain open

Record: [0662](../../0662-parallel-changed-member-deflate.md), implementing
decision 6 of [0652](../../0652-owner-decisions-for-the-third-wave.md) at the
`PreservationIndex::prepare` boundary. A managed source-backed publication now
computes a balance rule from the changed Deflate members, charges the admitted
wave's `CpuTasks`, reserves `Workers` and worker-state memory, and runs it on a
lazy private pool or an explicitly attached scoped facility. Results stay in
plan slots and emission remains in physical local and central-directory order,
so the 305 admitted OOXML fixtures publish the same SHA-256 at widths 1/2/4/8
and on the caller facility; 11 refusals retain their identity. The retained
publication measurements report 1.21×–1.52× paired median speedups on the
eligible real-fixture scenarios, with the same-window A/A floors beside every
row; the documented ordinary save remains publication-bound and carries no
execution context, so no saving is claimed for it. The implementation closes
the write row only. ADR 0031's other session rows are still open: the ZIP read
session remains eager and does not consume shared worker permits, and the CFB
bulk and source-backed multi-Part readers still lack the `IoConcurrency`,
`Workers`, `CpuTasks` and caller-facility bridge. Those gaps are stated rather
than counted as 0662 evidence. `performance_claim: none`; OLE2 and OOXML remain
active, ODF is deferred and iWork excluded.
[Record and limitations](0662-parallel-changed-member-deflate.md);
[retained evidence](results/change-0662/README.md).

## For `GOAL_AUDIT.md`

## 0662 — bounded write parallelism lands with deterministic bytes and an explicit ADR 0031 boundary

Record: [0662](0662-parallel-changed-member-deflate.md). The change follows the
goal's order at the first compatible parallel boundary: it removes serial
Deflate work for independently regenerated preservation members while keeping
all source audits, output preflight, cancellation fences, typed refusals and
exact no-op handling in place. The wave threshold uses the measured remainder
after the largest member, with a higher floor when a publication must build a
pool; a width-one or below-threshold publication stays on the old loop and
builds no pool. The corpus differential, error-order test, repeated
determinism test, physical-order test and sink-untouched cancellation and
budget tests establish the correctness side; the real-fixture medians and
same-window floors scope the measured side. The goal's execution-context rule
is met for this write path through explicit worker and task budgets and an
optional caller facility. It is not silently generalized to the existing read
sessions: `IoConcurrency` and cross-session composition for ZIP, CFB and
source-backed reads remain a follow-up boundary under accepted ADR 0031. No
ordinary CRUD signature exposes a scheduler, no hidden global pool is created,
and no `unsafe` is added. `performance_claim: none`; OLE2 and OOXML remain
active, ODF is deferred and iWork excluded.
[Record](0662-parallel-changed-member-deflate.md);
[retained evidence](results/change-0662/README.md).

## For `REPORT.md`

## 0662 — changed-member Deflate is a bounded publication wave, with savings scoped to the route that carries the context

The preservation writer can now compress regenerated Deflate members in a
bounded wave before writing the first byte. The balance rule is measured at the
publication boundary: warm sessions use an 8 KiB remainder floor, pool-building
sessions use 64 KiB, and a caller's per-task floor can narrow the granted width.
The ZIP sweep establishes the crossover and the largest-member bound; the OPC
sweep captures the cold-pool cost and caller-facility path; real fixtures report
1.21×–1.52× paired median end-to-end gains on eligible multi-member scenarios.
All outputs stay byte-identical at widths 1/2/4/8 and after seventeen repeated
width-eight publications. The 2.4% width-two regression on the one-dominant
`no_drawing_patriarch.xlsx` fixture is retained against its 0.35% A/A floor and
is below the 5% review trigger; no averaging hides it. These measurements are
publication measurements for managed source-backed overlays. The ordinary
save, eager writer, unmanaged package, streaming writer and replay route carry
no session and are unchanged, so the record makes no claim for them. The
implementation's accepted ADR 0031 scope is the write wave; the ZIP/CFB/read
session budget-composition rows are not included in these numbers and remain
open. `performance_claim: none`; no claim-registry entry is made.
[Record](0662-parallel-changed-member-deflate.md);
[retained evidence](results/change-0662/README.md).

## For `ADR_COMPLIANCE.md`

## 0662 — the write boundary satisfies decision 6; accepted ADR 0031's read rows are named as unfinished work

Record: [0662](0662-parallel-changed-member-deflate.md), authority [0652](0652-owner-decisions-for-the-third-wave.md)
decision 6. The production change preserves crate direction: `soapberry-zip`
defines its own scoped-worker trait, `litchi-core` defines the runtime-neutral
trait and budgets, and `litchi-opc` bridges them privately on the managed
source-backed publication path. The write wave reserves before compression,
never emits before every regenerated member is prepared, returns errors in
serial plan order, keeps cancellation cooperative and typed, and uses no global
pool or `unsafe`. Ordinary CRUD signatures remain free of archive, executor and
lock types. This record does not claim full implementation of every paragraph
of ADR 0031: `Resource::IoConcurrency` has not been added, the existing ZIP
read session still creates its pool eagerly, and the CFB and source-backed read
sessions do not yet share `Workers`/`CpuTasks`/caller-facility admission. The
remaining read-session requirements and their three ADR verification tests are
therefore explicitly outside 0662's decision-6 write scope, rather than being
described as closed by the new deflate tests. `performance_claim: none`.
[Record](0662-parallel-changed-member-deflate.md);
[retained evidence](results/change-0662/README.md).
