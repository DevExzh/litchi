# 0362ug — Edit construction cost audit

Record 0351 section 2.1 item 8 requires `Package::edit*()` to be O(1) in
package size. Section 2.6 item 3 additionally prohibits copying part payloads
at edit entry. Passing the public-signature inventory does not establish
either property. The final sweep found constructors that still materialize
source state before any operation is staged. These are required corrections,
not accepted performance exceptions.

This audit starts from assembled source `7559f44aa975254bd96f6e28eee17366b0ebeb6a`.
ODG and ODB/ODM were assembled at `8779d059ad`. Reviewed DOC and the ODS
resource-admission successor are assembled at `3676e21cfa`; reviewed ODP
entry/admission follows at `d316d04291`, and reviewed XLS at `50d4f4b88d`. Other rows below
remain open; this is source integration, not final performance acceptance.

## Confirmed scopes

| Scope | Entry cost found | Correction record and state |
| --- | --- | --- |
| ODG | Primary edit copied content and styles XML, including unused copies for pure joined package replacement. | 0362utf, reviewed and merged at `c3f56b7a9d`; lazy part state retains the source package and materializes on mutation. |
| ODB, ODM | ODB copied content XML; ODM title edits read raw meta.xml and copied projected title strings. | 0362ugb and 0362ugb-a, reviewed and merged at `8779d059ad`; borrowed XML/title state and deferred changed-title publication. |
| ODP | Package edit builds a complete mutable presentation and copies validated slides/source fragments; flat edit clones the slide vector. | 0362uga, reviewed and validated correction merged at `d316d04291`. The already-lazy tail-append entry is not this finding. |
| ODS | Source cells construct protection state; DDE and sheet-metadata constructors load source projections; package resource-count validation enumerates archive members; flat edits allocate per-sheet staging. | 0362ugc merged at `f144cd6029`; reviewed and validated 0362ugca merged at `3676e21cfa`, removing the added open/edit census while retaining admission at staging, commit and replay. |
| DOC | Primary/profiled and scoped families serialize/copy the source at entry; paragraph source-backed entry scans the source prefix. | 0362ugd and 0362ugd-a, reviewed and validated source/consumer correction merged at `cf7def80df`. |
| XLS | Primary and source-backed cells copy/index source state; scoped capture serializes the archive for comments, visibility, objects, names, charts and pivots. | 0362uge, reviewed and validated source/freshness correction merged at `50d4f4b88d`; the additional property/feature scope below remains required. |
| XLS properties and features | Metadata parses/copies properties, property-set entry captures the whole CFB, and worksheet features parse/copy the Workbook stream. | 0362ugea, deferred source and signed exact no-op correction under validation; facade refusal mapping and cached-source fencing are under review. |
| XLSX source-backed | Worksheet and defined-name/cell constructors load and decode selected parts or workbook state before staging. | 0362ugf/0362ugs, reviewed composed correction and planning-guard consumer merged at `3d129677ce`; package/consumer validation is pinned to the exact sources in 0362ugs. |
| XLSX pivot | Pivot cache/table/data entries load graphs and clone semantic drafts before an operation. | 0362ugn, reviewed composed correction and validation receipt 0362ugs merged at `3d129677ce`. |
| XLSX graph families | Metadata, calculation properties, connections, custom data, external links, data model and XML maps load selected graphs or snapshots at entry. | 0362ugo, reviewed composed correction and validation receipt 0362ugs merged at `3d129677ce`. |
| XLSX object families | Data-type icons, OLE objects, page breaks, revisions, rich-value refresh and surveys parse or copy selected source state at entry. | 0362ugp, reviewed composed correction and validation receipt 0362ugs merged at `3d129677ce`. |
| XLSX catalogs | Slicers and timelines census cache/worksheet parts; task panes load the web graph; smart tags parse selected worksheet XML. | 0362ugq, reviewed composed correction and validation receipt 0362ugs merged at `3d129677ce`. |
| Pages | Body-table constructors rebuild rooted inventories; chart data, drawable order, headers/footers and footnotes decode or copy selected/source graph state. | 0362ugg/0362ugr, reviewed and validated correction merged at `9ea602e0e1`. |
| Keynote | Text/name selection can decode a whole show; chart data and table/chart constructors decode selected payloads or scan the graph. | 0362ugh/0362ugr, reviewed and validated correction merged at `9ea602e0e1`. |
| Numbers | Title/image selection scans rooted graph state; chart and merge constructors retain decoded data/regions. | 0362ugi/0362ugr, reviewed and validated correction merged at `9ea602e0e1`. |
| PPT | Primary and metadata edits serialize and parse the whole CFB. `scoped::Base::capture` copies `to_encrypted_bytes()` into a new shared buffer. | 0362ugj, reviewed and validated correction merged at `127c26d59f`; standalone consumers and the ODG adapter compile, with corrected probe bindings. |
| PPTX | In-memory `Package::edit` captures all slides, notes and name indexes and fingerprints parts; opened edit construction also clones the slide vector and name index. | 0362ugk, reviewed and validated correction merged at `a3eb017da2`; remaining benchmark getter fixed under 0362ugka. |
| DOCX | Ordinary and source-backed document entry parses/indexes main XML; Ink scans package parts and bindings; OPC edit clones graph metadata; several source-backed document families capture source projections at entry. | 0362ugl/0362ugla/0362uglb, reviewed and validated correction merged at `5ce9829784`, including signed empty-property no-op preservation and all fuzz consumers. |
| XLSB | Primary and cell-value constructors serialize source before-images; several owner families read/clone selected graphs, and a cold `From<OpcPackage>` projection can materialize through generic staged entry. | 0362ugm, reviewed and validated correction merged at `0d35a83a31`. |

PPT's accepted 0362j work repairs replay and first-commit digest cost. It
does not authorize serialization at edit construction; the new finding is
separate. Cheap package `Arc` clones and cached immutable handles are not
findings. RTF's opened source tail-append entry remains cheap; its committed
deferred owner rematerialization is explicitly covered by 0359a. Likewise,
0358a's explicit rematerialization of a previously committed deferred ODS
owner is distinct from eager work on pristine source edit entry.

The remaining ODF owner audit was checked directly against the `8779d059ad`
tree. ODC and ODI facade edits clone flat snapshot Arcs; OTH and formula edits
retain a source reference and empty staging vectors. Their package snapshots,
and ODT's archive data and prepared index, are Arc-backed. Common metadata
entry only clones the owner. No source-sized edit-entry scan or payload copy
was found in these routes. This covers the separate ODC/ODI/OTH/formula crates
as well as ODT; the auditor's earlier assumption that they all route through
ODT was corrected before this disposition was recorded.

OLE-common object Package and OGraph Package entry retain shared source state.
The bounded DataSpaces, smart-tag, toolbar and similar codec Snapshot editors
are separately scoped by 0359; model/draft cloning there is not an archive
Package entry or a newly copied archive-part snapshot. Those observations do
not expand the Package O(1) finding to unrelated raw codec drafts.

## Review and validation boundary

Corrections retain immutable source handles, selectors and options until an
operation actually requires the source data. They must preserve source and
signature admission, typed budgets before retention, exact no-ops, patch
replay and inverse behavior, bounded deferred output and every existing
capability. Moving work into `open()` or hiding lazy failure behind a panic
does not satisfy this contract. A necessary fallible-accessor change migrates
every consumer in the same commit.

ODG's final source `05c15847c1a4ffc3f55e50c9ecc7609c9b338e0a` passes 23
library tests, 155 integration tests and one doctest, with strict checks;
0362utf records exact external logs. Independent reviews cover joined
composition, canonical effects, projected style dependencies, replacement
policies, optional styles and lazy XML lifetime. The merged correction
retains the complete return/correction history through receipt `0dca7284f9`.

ODB/ODM source `c44e1c57968cc1ddf227e814c85571b905c422bf` passes 106/58
regular tests respectively, one doctest each, and strict Clippy. Record
0362ugb-a labels its external receipts as reduced summaries of terminal
output, not raw command logs. Review confirms the borrowed state and
semantic behavior. ODB helper paths may materialize before discovering a
self-move or validation failure; construction/read-only/discard stay borrowed.
ODM raw meta.xml failures move from construction to a changed-title commit.

ODS source `2e2f902179d8ea107d05d3805542e89d9a42f27a` is independently
reviewed and passes 1,767 tests across 119 successful result blocks, including
702 library tests. Strict all-target/all-feature Clippy also passes. Both
final commands use the repository's pinned Rust 1.95 and tracked root Clippy
configuration. The raw logs are
`(deleted-coord-artifact)/ods-full-tests-pinned195.log`
(182,062 bytes, SHA-256
`81f35af676d3546b6f94f7b9253b72e3f5be9026677ace21a288855afd3b9bba`)
and `ods-all-targets-clippy-pinned195.log` beside it (729 bytes, SHA-256
`72fd3add599ba9f7e2557891ec1a07f579ac187eeb2ef1d0717681496079782d`).
The coordinator verified both hashes, every result block and actual root
configuration against Git. A preceding fixture expansion had excluded root
configuration and selected Rust 1.98; that run is historical evidence only.
All 37 required fixture paths were already tracked; no fixture was added.

A subsequent coordinator call-path review returned the resource-count cache
for correction as 0362ugca. `Archive::from_family_package` now calls `files()`
and counts filtered names during open, adding the complete name allocation and
census there. Passing tests and a cheap edit constructor do not justify that
cost relocation. The successor must derive the fact during an already-required
traversal or defer admission until first use, preserving limits before
retention. The 0362ugc test receipts remain valid for their exact source; no
native/performance acceptance is claimed for the returned census change.

## Follow-up review checkpoint, 2026-10-24

The ordinary XLSX audit is pinned to `8bf4833ae5`. It found the four
additional scopes above; retaining an Arc does not make a source-sized
catalog scan O(1). The pivot, graph and object corrections use separate
sparse worktrees with the actual root toolchain, Cargo and Clippy files
verified against Git. Necessary bounded caller-selector ownership is allowed;
source-name lookup and ambiguity scans remain deferred until first use.

ODS successor `8507ec866e` is independently clear at source review: the
resource census stays cold through open/edit/drop/handoff, and first staging,
commit and replay retain admission before candidate retention/publication.
The pinned rerun passes 1,768 tests and strict all-target/all-feature Clippy on
`1fdfcdcdba`; the only source delta after review removes redundant
`#[must_use]`. Receipt `daf6a85bd4` is merged at `3676e21cfa`. ODP `363a2497f9` likewise clears source review after moving security admission
before root/PI checks and covering the empty-splice boundary. The final run
passes 448 tests across 38 suites, strict all-target/all-feature Clippy and
the all-target/all-feature facade check. The coordinator verified all five
raw-log identities and root configuration. Receipt `88c77dfe11` distinguishes
the 20G memory cap from the 20,000,000,000-byte target cap and records the
unrelated formatting-only XLSX working-tree overlay during consumer checks.
Only committed ODP source was merged; that unattributed overlay stays in its
leaf. The previous 447-test receipt remains historical.

DOC source `251f33f1ce` now clears independent review and is merged at
`cf7def80df`. Failed profiled opens retain their observers; strict and fallback
paths retain the requested protection policy; ordinary edits remain Send/Sync
while profiled callbacks may borrow non-Send state. Every affected facade and
profiling consumer is migrated. The coordinator verified 1,445 passing default
tests across 50 result blocks, 1,223 profiled tests, both strict Clippy checks,
the DOC facade check and both standalone profiling checks. Record 0362ugd-a
identifies the exact clean source and tested two-file overlay; its hash matches
the committed final consumer delta. The original 1,222-test receipt remains
historical. No native timing claim is made.

XLS source `b8ed2d46bb` now clears review and is merged with receipt
`8c2b79f7c5` at `50d4f4b88d`. Independent content-matching replay retains the
recipient's deferred source fence. Both admission and later primary guards
promote in-flight composed-target errors to the original typed SourceChanged
when appropriate. Chart/pivot failures preserve retry state. The coordinator
verified 1,870 passing tests across 100 suites, strict Clippy, all-feature
facade checks and four metadata consumer tests. The one existing ignored test
writes an evidence workbook. Focused owner/source tests and the deterministic
admission regression also pass; raw logs and root configuration were verified.
Published-example doctests remain part of the composed validation below.

Other returns remain required before merging their leaves. Pages must
retain cumulative wire budgets, admit owned selector/text memory before
retention, avoid repeated payload copies, and migrate standalone fuzz
consumers. Numbers' named-selector scans are deferred in `86c762d2a7`, but review
returned uncharged selector retention/first-use traversal and stale standalone
fuzz consumers. PPT's native projection fence and deterministic in-flight
source-change regression clear review at `2f24654376`; runtime validation
remains pending. PPTX and XLSB source reviews are
clear after their final consumer/documentation corrections; Cargo validation
on those exact sources remains pending.

Keynote's remaining IWA failure was reproduced on clean baseline
`c3f56b7a9d`. The fixture used equal per-entry and aggregate limits but
expected the aggregate error for a replacement exceeding both. The narrow
repair raises only the per-entry limit. The repaired full IWA log contains
1,787 passing tests across 20 result blocks, including 1,696 unit tests and
8 doctests. Its leaf receipt has been returned for precise tested-overlay provenance.
The worker must also remove unrelated standalone lockfile upgrades and admit
owned selector memory before copying arbitrary names. Merge remains pending.

The reviewed glossary policy still parses as 55 profiles after these merges.
This is neither a fresh rustdoc inventory nor the 53-check gate. Complete
workspace Clippy on earlier source `8225b82dde` remains source-pinned evidence
and does not cover the subsequent constructors. Native qualification must
use the final assembled candidate. Comprehensive redaction audit 22 is the
only deferral; its content-free error and diagnostic rules still apply.

The source-backed XLSX packet at `02684ae108` now passes 2,542 tests and
strict Clippy after the slicer failure was reproduced on baseline `7559f44aa9`
and repaired. Ordinary slicer/timeline/task-pane/smart-tag source is committed
at `d14748b233` for separate review and a fresh validation run. The pivot
correction at `a99de1d698` is under source review; graph and object corrections
remain in progress. These later commits do not inherit acceptance from the
earlier source-backed test receipt.

## Composed validation assemblies, 2026-10-24

The coordinator assembled reviewed Pages, Numbers and Keynote corrections
with integration `164b0508db` in the existing Pages leaf. Initial composition
is `3ae5fdf0ff`; subsequent compile/test corrections retain the shared Footnote
Arc, fix a qualified test name and a merge-region fixture, and complete the
Keynote selector-kind mapping. Independent review confirms the Arc correction
has no payload-copy or failure-state regression. Full combined validation,
standalone consumers and doctests are in progress under record 0362ugr; those
corrections are not yet merged into main integration.

The XLSX assembly at `7ed6fb0ca7` includes integration through XLS
`50d4f4b88d`, source/catalog families, initial-batch deferred staging, pivot
`83a1b6a090`, graph `672b7498e8` and object `2ea78edd48` corrections. Nested
pivot validation now receives effective edit limits and preserves the typed
crate error; its new public regression distinguishes the 4,096-byte nested
part from earlier candidate serialization. Source reviews are clear and
combined Cargo/consumer/doctest validation is assigned under record 0362ugs.
The sparse checkout's omitted tracked `.cargo/config.toml` was materialized;
all actual root configuration now matches the composed source before builds.
The omitted file contains the lint alias, not compiler flags.

Reviewed ODF shared-triad tooling `7f85a0dd5d` is included in that assembly so
its Rust probe will be compiled in the all-feature facade checks. The
coordinator's three ODF unittest modules pass 18 tests; external log
`(deleted-coord-artifact)/tests.log`
is 117 bytes, SHA-256
`f10ec9ad7e06e8c5f41a75ab1f59b65505def0a0a9d05a795a2c958278cdb60e`.
Review covers normalized manifest/ZIP proofs, allocator observation boundaries,
actual source/config closure, per-endpoint toolchain context, journal ordering
and fresh outputs. No actual build or native measurement is credited by these
Python checks. The separate ODG protocol's source-byte, effective-environment
and final target-limit corrections now clear scoped review at `7c7d7a54fb`.
The coordinator independently ran its 26 Python tests successfully and merged
the tooling at `a41b437ca8`, removing only a trailing blank line in the Python
bindings file during merge. Its Rust adapters still require real compilation;
this is tooling integration, with no native qualification claim.

PPTX source `c10d6c6539`, comment correction `fa63bdd22e` and receipt
`e59c72184d` merge at `a3eb017da2`. Candidate reopen is attempted on every
route; only recompressing read-limit failures use the bounded fallback.
Output limits remain separate from execution-memory admission. Independent
review confirms lazy sharing, retry state and serialized-patch limit checks.
The coordinator verified 1,837 passing tests across 113 result blocks,
48 doctests, strict Clippy and 11 facade tests, with all raw-log identities
and root configuration. One existing regular test and two doctests remain
ignored. Record 0362ugk identifies an earlier uncapped diagnostic separately;
it is not accepted validation. No native timing claim is made.

The composed iWork suites, 96 explicit doctests and strict four-crate Clippy
pass at their recorded source boundaries through `4c323ae060`; facade and
standalone-consumer checks continue. A scoped review of changed diagnostics
returned the Pages table-name Edit's authored-name Debug field, and review of
its lazy state found that failed resolution drops the charged wire budget.
Both require correction and focused validation before merge. This is not an
expansion of deferred audit 22. Existing typed refusal enums remain permitted
by 0351 §2.4.2.

The four owned build slots now cover composed iWork, composed XLSX, XLSB and
DOCX. PPT runtime validation remains queued. Native endpoints stay unresolved
until final source freeze; no required engineering work is waived.

## Validation follow-up, 2026-10-24

Pages correction `57b14d3f2a` clears scoped review of all three new persistent
lazy ledgers (table name, appearance and merge). Failed resolution retains the
charged ledger and the next attempt reuses it; typed retry tests cover the
operation-work ceilings. The modified name Edit's Debug now reports presence
and byte count only. The coordinator verified the final Pages log: 719 tests
across 61 result blocks, with no failures or ignored tests, and actual root
configuration matching the source. Strict Pages Clippy also passes. Remaining
combined iWork consumers and the final receipt are still in progress.

DOCX validation corrections clear review at `d85c59cf91`. Signature admission
uses the retained physical source before the effective-topology branch, so a
chained deferred edit does not fail by demanding a materialized package.
Formatting consumers force their deferred projections; typed errors,
no-op/signature behavior, retry state and bounded paragraph routes remain.
Full runtime validation continues. XLSB validation corrections clear independent
review at `51c8628446`; their runtime/consumer checks continue. XLSX corrections
through `8c39f97fe1`
clear review; the later `d486a2490f` only forces the first data-model read in
three refusal tests. Its combined runtime validation continues.

The PPT validation leaf is composed with integration `d0e1782f9b` at
`7ae7e9144b`, retaining the reviewed native source-fence correction. It awaits
the next build slot. The final gate runner's first post-exit successor
`7f8b1a01e5` passes its verifier but is returned: a detached `setsid()` helper
can escape the current process-group lookup and write the target after a pass.
An independent controlled reproduction establishes the gap. Ownership tracking
and its detached-child regression must be corrected before final gate use.

Gate successor `c196274831` closes that return and is merged at `72d106cfac`.
The runner captures PID/start-time identities during a step and adopts new
subreaper children without relying on cwd or command-line paths. The final
regression covers `setsid`, `chdir("/")` and a delayed write through an
inherited target descriptor; the predecessor fails that case. The coordinator
independently verified the regression and byte-exact patch reproduction.
Attempt-specific logs preserve retry provenance. This clears the runner for
execution; it does not count as running the 53-step gate.

The composed iWork ops check passes at `5ac9b25891`. That older non-iWork
source boundary requires direct XLS/PPT metadata property access in its bench
example. The actual merged XLS getter remains infallible; the queued PPT
packet must migrate its own getter call when it merges. Standalone fuzz validation found one missed
fallible Keynote `position()` result; `a0bf134246` now observes the error and
returns before further work. The rerun passes with only the three pre-existing
unused media constants. The final iWork harness check and receipt remain open.


## Reviewed assembly merges, 2026-10-24

Composed iWork source `a0bf134246` and receipt `4441c86bbe` merge at
`9ea602e0e1`. The coordinator independently verified all successful and
reported diagnostic logs, actual root configuration, three lock identities,
and the minimal harness lock delta. Final package runs total 4,695 passing
regular tests, with 96 passing doctests and 50 selected facade tests. Strict
checks, operations targets, standalone fuzz and unified-batch iWork checks
pass at their documented source boundaries. Existing ignored tests and the
three unrelated fuzz media constants remain explicitly recorded. This is
source/correctness integration; native qualification remains open.

XLSB source `51c8628446`, documentation correction `c7aecdb536` and receipt
`16ba8da04f` merge at `0d35a83a31`. The coordinator verified the five closed
logs: 1,355 package tests including 42 doctests, strict all-target/all-feature
Clippy over XLSB/ops/facade/Python, 126 operations tests, 23 facade tests and
nine standalone CRUD harness tests. All actual root configuration and both
locks match their recorded identities. The final target was 12,379,273,678
bytes, below the decimal 20 GB cap. The receipt now distinguishes the r7
pre-footer hash and the final closed log, and identifies the documentation-only
slicer/timeline overlay that overlapped the standalone run; exact compiler-read
timing was not captured. No native performance claim is made.

The combined CLI/ops/facade/Python strict check first found a missing PPTX
benchmark migration at `0d35a83a31`; the gate monitor also hit a vanished-PID
race. Correction 0362ugka propagates `slides()` errors, and runner successor
`f9cf4aa6d5` narrowly handles `ProcessLookupError`. The rerun passes at
`0f007c4d74` in 27.094 seconds with all targets/features, a clean source,
no remaining owned processes, and stable target size 4,706,526,684 bytes.
Its reduced source/command/log/plan/state receipt is in 0362ugka. This is
one selected check, not the final 53-step gate. The released coordinator slot
is assigned to the XLS follow-up alongside PPT, XLSX and DOCX. Final source
freeze, all-68 inventory, full gate, corpus and native qualification remain
required.


A subsequent check of the actual XLS metadata getter exposed an omitted
constructor path: `metadata::Edit::new` still constructs a property-set editor,
parses `SummaryProperties` and clones the initial values before returning.
The earlier XLS receipt establishes its tested source behavior, but does not
close this entry-cost gap. Required follow-up 0362ugea is assigned in a fresh
leaf for metadata/property-set lazy state and every consumer. The bounded
review also found `sheet_features::Edit::new` parsing globals and copying the
complete Workbook stream. These three families are the finite follow-up scope;
other public Package edit entries checked are already lazy. This is engineering
work before source freeze, not an accepted exception. The current XLS getter
is infallible; its eventual migration must follow the actual corrected API.


## XLSX composed merge, 2026-10-24

Source `784d0ea83e`, final receipt `7201866b32` and the composed source-backed,
catalog, pivot, graph and object corrections merge at `3d129677ce`. Independent
review confirms lazy admission, retained budgets, source/signature behavior,
retry atomicity and the migrated consumers. The coordinator verified all twelve
closed log identities, actual root configuration, used lockfiles and four
consumer/lint overlay hashes plus the final planning-guard overlay.

The recorded package run has 2,558 passing tests across 98 result blocks and
36 passing doctests with no ignored tests. Strict Clippy covers the documented
implementation lint overlay; it does not relabel the earlier clean-source test
run. Operations, standalone performance consumers and the fuzz consumer compile.
The guard's valid, late-validator and late-raw diagnostic routes prove exact
no-op/source preservation or typed first-use refusal and retry preservation.
Unknown cell attributes remain supported and preserved; the stale validator
fixture now uses an explicitly unsupported child element. Debug timing fields
from these diagnostic runs are excluded from performance evidence.

The only merge conflict was formatting of the XLS/PPT metadata benchmark.
Resolution preserves their currently infallible property getters and the
already-fallible DOC getter. Pending PPT and XLS-property source changes must
migrate their own getters when they merge. No native qualification is claimed.


## DOCX and PPT merges, 2026-10-24

DOCX final source `0cf1a17174` and receipt `3a6c56958c` merge at `5ce9829784`.
Independent review covers source/limit precedence, inverse/fidelity assertions,
lazy projection access, signed valid-target exact no-ops, invalid-target errors
and complete fuzz publication migration. The coordinator verified ten captured
logs: 2,355 package tests, 170 doctests (31 ignored), 739 operations tests,
23 facade tests, focused signed regression, strict checks and standalone
consumers. Record 0362uglb distinguishes the exact clean pins and the two-file
fuzz overlay. Merge resolution only normalizes cfg feature order and receipt
wording/whitespace, adds verified missing identities and preserves every source
change. No native qualification is claimed.

PPT implementation/operations source `ca125d2ba1` and final consumer/binding
receipt `70a9b57622` merge at `127c26d59f`. The coordinator verified 1,568 package
tests (three ignored), 22 doctests (eight ignored), 71 selected operations
tests, strict Clippy, all-target consumer checks and standalone compilation.
The shared OLE benchmark explicitly handles the now-fallible PPT property
getter; XLS remains infallible until its separate follow-up merges. The ODG
candidate adapter compiles through the separately hashed external overlay.
All three changed probe identities are refreshed, and 26 build-free ODG tests
pass. These results qualify compilation and correctness checks, not native
performance.

The composed strict consumer check passes at `ecf598b637` after a private
ODF probe identity grouping closes its fourteen-argument lint failure (0362ugsa).
This pin precedes the DOCX/PPT merges and is not the final gate. The remaining
known source returns are XLS changed-publication freshness under 0362ugea and
the PPTX refusal-kind classifier. Final full gates, all-68 inventory, corpus,
glossary and native qualification remain required.


The subsequent combined strict check also passes at
`99c875ca41d7fa4a6185b2fff122e4499b6c0d88`, including the DOCX and PPT
merges: all-target/all-feature CLI, operations, facade and Python Clippy
exits zero in 145.777 seconds. Its source stays clean, three post-exit
observations find no owned builders, and the target stabilizes at
4,706,726,912 bytes. The [reduced receipt](results/0362ug/composed-consumer-check.json)
binds source, command, locks, policy, plan/state and logs. This selected step
remains distinct from the required full 53-step gate.

PPTX refusal-kind correction 0362ugt merges at `c1eda2fdf6`. `FinalSlide`
now follows its existing `InvalidValue` kind; signed/protected refusals retain
`document_protected`. Operation and reason text no longer override the code.
Classifier, slide-removal, capability and registry tests pass, as does the
strict PPTX-only operations check. Both generated consumers and the one-line
capability-table digest are synchronized. The coordinator verified all seven
logs/receipts and root configuration, then independently reran both static
OfficeCLI table checks on the composed tree. The final actual corpus replay
remains part of final acceptance. No unrelated DOC behavior changes.
