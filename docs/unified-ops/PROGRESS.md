# Unified operations API — progress log

Driven by `docs/UNIFIED_OPS_GOAL.md`; specification `docs/UNIFIED_OPS_API_DESIGN.md`.

## Records

| No. | Title | Status |
|-----|-------|--------|
| 0036 | OPC source-limit kinds and canonical errors | physical kind/error fields migrated; broader limits policy pending |
| 0035 | OPC/OOXML source-limit vocabulary | shared physical profile and builder renamed; broader limits policy pending |
| 0034 | Effective source DOCX text streaming | retained main/donor output integrated; specialized routes pending |
| 0033 | Effective source DOCX metadata reads | shared property graph/decoder integrated; remaining specialized routes pending |
| 0032 | Typed primary DOCX transaction limits | typed kind/count fields migrated; configurable limits and broader policy pending |
| 0031 | Canonical source DOCX package-owning Edit/Commit | primary owner migrated; specialized families, limits and performance pending |
| 0030 | Immutable retained DOCX semantic apply | main-document apply integrated; independent managed donor snapshots refused |
| 0029 | Source DOCX effective owner and physical-operation boundary | effective reads/publication integrated; specialized routes pending |
| 0028 | Effective OPC owning conversion | explicit materialization integrated |
| 0027 | DOCX patch predicate and content-free Debug | primary patch migrated; broader diagnostics pending |
| 0026 | Owned reads and physical identity access for retained OPC candidates | effective read adapter foundation; integrated in 0029–0031 |
| 0025 | Immutable XML edits of retained effective OPC candidates | flat chained source-Part candidates; primary owner integrated in 0031 |
| 0024 | Revision-bound source XML edits over retained payloads | per-part chaining foundation; package overlay revisions pending |
| 0023 | Retained OPC prepared topology and shared ZIP appended members | source-overlay ownership foundation; chained commits integrated in 0031 |
| 0022 | Eager DOCX package-bound Edit/Commit and removal of mutating publication doors | eager primary commit migrated; source joined in 0031; history and full W2 pending |
| 0021 | Shared source-backed DOCX owners and borrowing stream publication | sharing prerequisite complete; primary package-bound triad integrated in 0031 |
| 0020 | Cheap eager DOCX package shares and immutable primary patch application | eager apply migrated; package-owning commit W2 work pending |
| 0019 | Owned eager DOCX read projection and shared physical graph | projection migrated; immutable package and commit W2 work pending |
| 0018 | DOCX primary edit entry and focused error/result names; minimal/full glossary bans | names migrated; package-owned commit/projection W2 work pending |
| 0017 | W0 rustdoc API glossary gate, scoped rules, six-type family checks and CI | complete; profiles for leaf migration waves pending |
| 0016 | Rebase onto current base; adopt API glossary as ADR 0034; preserve shared fingerprint memo | kickoff; W0 checker and leaf migrations pending |
| 0001 | Owner override: read + write for every claimed format; tiers OOXML → OLE → (RTF, ODF, iWork) | recorded in the goal (2026-10-04) |
| 0002 | Leaf write-capability inventory (deep: OOXML + OLE; brief: tier 3) | merged (`cfe7681627`) |
| 0003 | M0a facade migration plan (flat `File`, `litchi::open`) | merged (`45a932ea2e`) |
| 0004 | ADR 0033 draft + `litchi-ops` scaffold + gate registration | merged (`af70966a6d`) |
| 0005 | M0b path/selector grammar in `litchi-core::selector` | merged (`9f8e49b181`) |
| 0006 | Owner accepts ADR 0033 | done |
| 0007 | Coordinator defaults for 0003's open questions | adopted, awaiting owner |
| M0a | Steps S1–S10 on `uo/m0a` (plan 0003 §4) | merged (`1927763795`); S10 gate passes with E1/E2 awaiting owner (F1, F2) |
| 0008 | Diagnosis: source-backed DOCX edits refuse every Word-authored fixture (MCE) — M2 blocker | merged (`e9c3868109`); fix staged as 0011 |
| 0009 | Design: editing a shared snapshot (R4) — M2 prerequisite | done (`9b3cab60f6`); staged leaf plan S0–S5 |
| 0010 | M1 unified-read plan (node view, path resolution, query, view modes, capabilities) | done (`31bd660497`); steps R1–R11 with leaf waves L1/L2 |
| M1 | Steps on `uo/m1` (plan 0010 §10); merges after M0a | in progress |
| 0011 | DOCX MCE admission, stages 1–3 of 0008 | merged into `uo/m0a` (`53f938e83c`); 19/19 Word fixtures editable source-backed; eager owner gains the same refusals (breaking) |
| 0013 | Signature policy: `Signed::unsign`, builder refusals, pptx/xlsb silent-strip fixes (M0a S6b) | merged into `uo/m0a` (`f30c521c90`) |
| 0015 | M2 write-kernel plan (verbs, registry, per-format handlers, batch executor, leaf prerequisites, save/create) | done (`de8d557c56`); steps W1–W10 and leaf prerequisites L-*-W* |
| M2 | Steps on `uo/m2` (plan 0015 §11); merges after M0a and M1 | in progress |
| 0014 | Implicit `unsign()` sites: pptx and shared ooxml-common (0014a), xlsx/docx and XLSB task panes/Ribbon (0014b); 0013 F1, F1b, F2 | merged into `uo/m0a`; silent strips now refuse until explicit `unsign` (breaking) |
| 0012 | litchi-opc: cheap `Clone` for `SourceBackedPackage` (0009 stage S0) | merged into `uo/m0a` (`0cf1f43403`); clone 6 reference counts, no allocation; open unchanged |

## Awaiting owner confirmation

- M0a S10 (`docs/unified-ops/s10-measurement.md` §9). M0a landed on the strength of two explanations. If either is rejected, the chosen alternative becomes a harness follow-up; M0a is not reverted.
  - F1: accept E1. Seven rows over their trigger time one call in a fresh child process, so they measure the harness executable. Standalone probes give the same first-call time at both revisions (XLSX open 187.6 vs 187.9 µs); Callgrind shows +0.75% instructions. Alternatives: an untimed warm-up call in the child (new route names), or standalone binaries for first calls.
  - F2: accept E2. The RTF open row depends on the allocator state the untimed projection leaves; with malloc trimming off, the unified open is 0.98× the pre-M0a one. Alternatives: fixed malloc tunables for RTF children, or bare-loop variants.

- Record 0015 decisions D1–D19 and open questions Q1–Q14 (provisional defaults). Highlights:
  - Write verbs take `&self` and return a new `File`; `edit` returns a `Commit`.
  - Batches run sequentially, grouped into same-family runs, and the batch result must equal its items applied one at a time.
  - Writes refuse invalid values and unknown keys, with suggestions. OfficeCLI's apply-the-rest-and-warn behaviour is only a command-channel option.
  - After `Signed::unsign`, a save writes the unsigned copy. Decrypted files save re-encrypted when given the password, or refuse; plaintext needs `into_plaintext()`.
  - XLSX moves to the eager engine on its first write; PPTX text-only runs stay on the source-backed editor.

- Record 0010 open questions (§11.3), the two that most need the owner:
  - Q3: regex values (`~= r"…"`) are refused with `unsupported_capability` until the owner approves adding the `regex` crate.
  - Q6: sheet names match case-insensitively in every workbook format. This changes XLSB and ODS.
  - The other eleven defaults are listed in §11.3.

- Record 0009 (R4): a shared handle shares the leaf package instead of moving it. Write verbs take `&self` and return a new handle; there are no consuming verbs on `File`. Option B (keep the reader and re-open) and option C (consuming verbs with a sole-owner fast path) are rejected.

- Record 0007 defaults: flat `File` (F2), `limit_exceeded`/`cancelled` codes (Q1), iWork → `unsupported_format` until tier 3 (Q2), `retain_source = true` (Q3), `Signed` wraps every signed input (Q4).

- `OpenOptions` shape: one struct with common knobs plus optional per-format sub-options.
- Verb names: `apply`; find/replace as `find=`/`replace=` props on `set`; `merge` = template merge.

## 0004 — log

- Drafted ADR 0033 (Proposed) and listed it in the ADR index; ADR 0024 inventory amended (66 packages / 245 declarations).
- Scaffolded inert `litchi-ops` (edge to `litchi-core` only); registered in `crate_boundaries.json` and `non_iwork_gate.py` (bulk 46 → 47); gate tests updated and green.
- Known: workspace `cargo fmt --all --check` fails on base-branch files (soapberry-zip, iwa-archive) unrelated to this program.
- Follow-up: the first iWork leaf edge from `litchi-ops` (tier 3) requires reclassifying it out of the bulk gate, like the facade.

## 0002 — log

- Code-verified inventory; facade is read-only for every format and drops/never holds the editable leaf owner for doc, ppt, rtf; xlsx/xlsb are held only in the non-editable lazy form.
- No `Locked`/`Signed` types exist. Changed commits silently strip signatures in pptx `opened::Transaction`, several xlsb families, odf-formula, ODS legacy mutators, ODP in-place mutators (ADR 0006 concern; fix in the signature-policy prerequisite).
- doc/ppt `Writer::set_property` never emitted; xls `Writer` writes no property streams. Feature-matrix overclaims: pptx hyperlink write, docx app.xml write.
- Stream order: T1 xlsx > docx > pptx > xlsb; T2 xls > ppt > doc; T3 per record. Cross-cutting prerequisites: facade atomic writer, signature/encryption policy, commit-shape adapter, editable owner retention in M0a.

## 0003 — log

- Plan for M0a: a flat `File` enum with one variant per format, plus `Locked`/`Signed`, each variant one `Arc` handle (16 bytes, Send + Sync). The editable owner is built at the first edit, not at open. Leaf constructors L-PPTX-1, L-XLSX-1, L-XLSB-1, L-DOC-1, L-XLS-1 and L-PPT-1 avoid indexing the container twice.
- Gate: `non_iwork_gate.py` forbids any edge from a bulk package to an iWork crate. So `litchi-ops` stays free of iWork until M6, and the gate must be redesigned first (R1).
- Findings against the design: F1 (§16.2's "mutable owners" for pptx/xlsx are read-only), F2 (flat enum), F14 (Python `FileFormat` drift, including Odb → Doc).

## 0005 — log

- `litchi-core::selector` grammar: paths (1-based → `Position`, `last()`, quoted segments), predicates with `and`/`or`/`not`, header-row predicates, CSS-like queries (`:contains`, `:has`, `:not`, combinators, unions, sheet scopes, ranges), story axes, typed `ParseError` with offset and code, explicit limits.
- 950 OfficeCLI paths and 132 selectors parse and round-trip. Typical paths and simple queries parse with zero allocations (~57 ns per path, ~120 ns per selector, release build, single machine). There are 30 recorded OfficeCLI deltas: element aliases and bare-key meaning move to M1 resolution, and the unscoped-mutation guard moves to the verb layer.
- Open: the `selector_grammar` fuzz target compiles but has not been run (no nightly toolchain or cargo-fuzz on the host).

## 0008 — log

- None of the 19 DOCX fixtures uses a content-changing MCE construct. 13 only declare the namespace (Word 2007 `xmlns:ve`) and 6 carry one root `mc:Ignorable`.
- The source-backed guard refuses them because it compares buffers, and the MCE codec rewrites any part that mentions the namespace URI. The guard also has a hole: a namespace URI written with a character escape slips past both the codec and the guard.
- The eager owner keeps every `mc` construct byte for byte, but it has no MCE or protection check. It accepts a wrong-target edit under body-level `AlternateContent`, accepts `MustUnderstand` documents, and publishes edits to 3 protected fixtures.
- `to_owned_package()` rebuilds the ZIP layout (M0a `SourceMode::Eager` uses it). The execution-context route refuses even reads of all 19 files.
- Reads and edits count paragraphs differently: reads include table cells, edits count only top-level body paragraphs. M1 has to choose for `/body/p[N]`.
- Recommended: a staged admission rule (declared-only, 13/19; then Ignorable-only, 19/19) applied to both owners and the execution-context route, with one protection policy. Run-level `AlternateContent` comes later with a post-commit check. No silent eager fallback.

## 0009 — log

- No leaf guard needs a unique owner. The OOXML publishers take `self` by API choice but only borrow inside, and every share of one opened package carries the same lineage token. The packages cannot be cloned only because their index, catalog and cache are owned inline.
- Probe: the xlsx shared-handle fallback in `Xlsx::into_editor` breaks exact no-ops on all 4 real-producer fixtures (0 of 10–13 ZIP members kept byte-identical) and makes output depend on whether a second handle exists. It must be retired before M2.
- No OOXML source-backed commit can be chained before publishing, so M2 publishes into memory and re-opens at most once per family boundary until overlay chaining (option D) lands.
- Staged leaf plan: S0 litchi-opc cheap `Clone`; S1 xlsx (ops only); S2 docx `&self` publishers; S3 pptx; S4 xlsb non-mutating apply; S5 OLE needs nothing for sharing.

## 0010 — log

- M1 plan:
  - A borrowed node view serialized straight to OfficeCLI's `DocumentNode` JSON, built from static per-element property tables (no per-node map).
  - Leaf-owned evictable indexes.
  - `/body/p[N]` is OfficeCLI's direct-body paragraph, counted on the MCE-visible view. M2 maps it to the edit transaction's numbering under 0011's admission check.
- The default source-backed owners answer little today: DOCX exposes only the main story, PPTX only slide text, XLSX only cells. About 35 core leaf read requests plus 4 correctness fixes are listed in §4.11 (L1 tier 1, L2 tier 2).
- Steps R1–R11. R1–R3 and L1 start before M0a lands, on `uo/m1`; R5 waits for S6, R8 for S9. M1 merges into trunk after M0a. 17 code findings (F-1…F-17) and 21 OfficeCLI deltas.

## M0a — log

- Landed as one merge of `uo/m0a` (`1927763795`). Step detail: `docs/unified-ops/m0a-log.md`.
- `litchi::open`, `open_bytes`, `open_with` and `open_prepared` return the flat, non-exhaustive `File`: 16 bytes, one `Arc` handle per format, plus `Locked` and `Signed`. Detection lives in `litchi_ops::detect`.
- One pointer-sized `Error` with an OfficeCLI-aligned `Code` and a typed `Source`.
- Deleted (S9, −27k lines):
  - the facade's `Document`, `Presentation`, `sheet::Workbook` and sheet openers;
  - the old Markdown adapter;
  - the `__migration` seam.
- The facade, examples, tests, Python bindings and perf harness all use `File` (S7, S8a–d).
- Breaking changes:
  - Changed commits that silently stripped a signature now refuse until `Signed::unsign` (0013, 0014a/b).
  - The eager DOCX owner refuses what the source-backed one refuses (0011); 19/19 Word fixtures are editable source-backed.
- `SourceBackedPackage` clones cheaply (0012).
- S10 gate: passes on the plan's terms.
  - 91/99 facade rows and all 47 kept predecessors pass the raw rule; the other 8 are E1/E2 (above).
  - R5: DOC and the PPT byte route keep about one input size alive after open, and `retain_source(false)` removes it.
  - Follow-ups F3 (drop the whole-child RSS tripwire, which measures the harness's `git` child) and F6 (captures off the shared tmpfs).
- Known failures inherited from the base, unchanged:
  - workspace fmt (iwa-archive, soapberry-zip);
  - 18 facade feature-subset sets;
  - the `keynote_chart_title` fuzz target;
  - rustdoc without features.

## 0015 — log

- M2 plan:
  - Verbs take `&self`: the sugar verbs return a `File`, `edit` returns a `Commit`, and `batch` is atomic or best-effort with per-item results.
  - The registry grows from M1's `PropSpec` and refuses rather than repairs.
  - Batches run sequentially, in runs, with barriers and fusion, and rebase only at family boundaries.
- Critical path: no tier-1 leaf can format existing content today, so the parity engines come first: L-DOCX-W5 (DOCX formatting), L-XLSX-W2 (style allocation) and L-PPTX-W3/W5. They follow 0009 S1–S3 (sharing) and the DOCX run, cell and story addressing.
- Leaf measurements (leaf APIs only):
  - Coalescing is essential: 1k chained XLSX commits took 0.18–12.6 s, against 0.5–7.8 ms in one transaction.
  - The XLSX source-backed cell editor admits 45 of 137 corpus workbooks; the eager route admits 132.
  - PPTX source-backed text edits are 7–18× cheaper than the eager transaction.
  - A DOCX reopen between runs costs 9.4 ms on 10k paragraphs.
  - An XLSX commit rewrites the whole sheet: 53.5 ms for one cell on a 75k-cell sheet.

## Awaiting owner decision

- ~~ADR 0033 acceptance~~ — accepted 2026-10-04 (record 0006).

## 0016 — restart and glossary adoption (2026-10-08)

- Rebased `feat/unified-ops-api` from `dbe66611f1` onto base `32ad64e320`;
  rebased tip before this record is `61630e5635`. Existing M0a/M0b work is
  retained. Historical hashes above are evidence references, not current tips.
- Adopted the owner-supplied glossary/default-policy draft as ADR 0034;
  preserved the existing accepted operations ADR 0033 and its citations.
- Resolved the OPC cheap-clone/fingerprint-memo conflict by sharing the base
  memo through package clones; added regression coverage for clone survival.
- W0 next: rustdoc JSON checker with scoped rules, tests and CI integration.
  Then DOCX → PPTX → XLSX shape waves with simultaneous consumer migrations,
  followed by the remaining format waves and M1/M2 completion. iWork requests
  are recorded for its owners; no iWork implementation changes here.
- No performance, parity, all-format CRUD or milestone-completion claim is
  made by this kickoff. Verification and remaining gates are in record 0016.

## 0017 — W0 glossary enforcement

- Added the rustdoc public-API checker, explicit feature-profile policy, real
  compiler fixtures and CI integration. W0 inspects `litchi-core` with empty
  bans/families; pending API migrations are not claimed conformant.
- Scoped bans cover public names, aliases/methods and public signature type
  names. Registered transaction modules require all six canonical type names.
- 19 tests and the current-tree W0 profile pass. Review findings and rustdoc
  limitations are recorded in 0017; unresolved external globs fail closed.
- Next: DOCX W2 migration with consumer updates and behavior gates. W1 iWork
  ownership and the original all-format CRUD goal remain unchanged.

## 0018 — DOCX canonical entry/error names

- Eager/source `Package::edit()` replaces `edit_document`; main-document
  errors/results now live at `document::{Error, Result}`. Existing edit,
  commit, patch and publication semantics are retained for the ownership step.
- Active consumers, examples, tests and current leaf docs migrated together;
  historical benchmark records keep their original source/API names.
- Minimal/full DOCX glossary profiles enforce removal without public aliases.
- Verification and outstanding W2 ownership work are recorded in 0018.

## 0019 — owned eager DOCX projection

- `Package::document()` returns owned `Document`, retaining a shared physical
  snapshot. It remains usable after package drop and stable across later edits.
- Existing eager mutation paths detach through copy-on-write; clean saves
  preserve the shared graph. Main-part payloads and paragraph indexes remain
  shared, with related-part resolution tied to the retained graph.
- Updated dependent Rust types, scoped style borrows, runnable documentation,
  glossary status and ownership/rollback/no-op regression coverage.
- This completes the eager projection shape, not the immutable package or
  package-owning commit contract. Evidence and remaining work are in 0019.

## 0020 — immutable eager DOCX patch application

- Eager `Package::apply(&patch)` returns a new package and replaces the removed
  mutating `apply_document_patch`. Active consumers and the glossary gate move
  together; the existing validation/publication engine stays private.
- Package clones share the physical graph, metadata and legacy writer state.
  Mutations detach state; failed saves retain edit intent on the affected clone.
- Regression coverage includes immutable apply/inverse/stale rejection, no-op
  sharing, writer/metadata isolation, signatures and encrypted-source policy.
- This is the publication prerequisite for package-bound `Edit`/`Commit`.
  Those shapes, source-backed package publication and writer removal remain
  required; evidence and limits are recorded in 0020.

## 0021 — shared source-backed DOCX publication

- Source-backed `Package` clones share source identity, lazy caches, admission,
  tolerance reports and execution reservations without reopening the archive.
- All 32 package stream publishers borrow their owner. Retained owners support
  repeated publication and retry after sink failure; source-version checks stay
  active. Cache reservations persist until the final retaining owner drops.
- Concurrent clone/no-op and metadata-report tests cover shared state; existing
  managed tests explicitly release the retained source before zero-budget checks.
- This enables retained source publication; package-bound `Edit`/`Commit`, source
  overlays and canonical immutable publication remain required. See record 0021.

## 0022 — eager DOCX package-bound commit

- Canonical `document::Edit` retains its complete eager source. Its commit
  validates and publishes a new `Package` with the patch and diagnostics triad.
- Removed eager `publish_document_edit`/`publish_document_commit`; migrated
  consumers and enforced removal in minimal/full glossary profiles.
- Explicit nominal `document::snapshot` types retain the raw XML engine for
  standalone/source/history callers until their owner migration. Root commits
  cannot carry only a snapshot or an optional owner.
- Tests cover retained related parts, source drop, forked edits, chaining,
  inverse, no-op sharing, later source writer changes and security refusals.
- Source-overlay owners, snapshot composition/history, full package lineage,
  limits and writer removal remain required. Evidence and limits are in 0022.

## 0023 — retained source-backed topology publication

- OPC `PreparedTopology` now cheaply clones its complete prepared candidate
  and reservations. Publication borrows it, enabling retry and repeated output.
- ZIP append plans retain shared generated-member handles; publication does
  not deep-copy prepared relationship XML, member names or appended payloads.
- Added metadata admissions and tests for concurrent clones, retry, source and
  transfer rechecks, candidate readback and final-owner reservation release.
- This is the retained topology foundation, not chained source-package commit.
  Logical overlay revisions and revision-bound XML proofs remain required;
  the original physical source proof is deliberately unchanged. See 0023.

## 0024 — revision-bound source XML edits

- `SourceXmlPart::begin_revision()` establishes a fresh logical edit base over
  the current validated payload, without reopening the ZIP or rescanning XML.
- Range proofs bind to that revision, physical source, part and content type;
  sibling and stale proofs are refused even when their bytes match.
- Physical replacement authority remains the original archive. Managed proofs
  retain payload admissions; sequential revisions do not retain an ancestor chain.
- Chained edits, inverse exact-artifact publication, source fences and managed
  limits are covered. Package-wide effective revisions and canonical source
  commits remain required. See record 0024.

## 0025 — immutable edits of retained effective candidates

- `PreparedTopology::edit_xml` applies a checked edit to effective source XML,
  returning a new flat candidate over the same physical archive.
- Prior additions, omissions, relationships and content types survive; logical
  no-ops preserve the current candidate, and exact inverses restore source bytes.
- Effective parts retain XML provenance. Unchanged metadata and transfer guards
  share ownership; old candidate vectors and intermediate payloads can release.
- Revisions enforce physical/transfer fences, aggregate limits, operation bounds,
  cancellation and managed admission. Swapped callback transactions are refused.
- Added/unproven XML targets, general topology application and canonical source
  DOCX commits remain open. Record 0025 covers evidence and performance limits.

## 0026 — owned effective Part and payload reads

- Retained candidates expose owned lazy Part views and payload handles, including
  unchanged candidates, plus effective iteration and main-document resolution.
- Added/removed Parts, effective relationships and changed payloads route through
  the captured catalog. Views survive owner drop and remain stable across edits.
- New captures recheck base/transfer sources and execution contexts; captured
  metadata and bytes remain immutable readable snapshots. Payloads retain their
  managed admissions and the complete candidate, without a detachable Arc.
- XML proof capture is shared with `edit_xml`. Narrow physical identity, artifact,
  limits and execution accessors avoid requiring a parallel raw package owner.
- Canonical source DOCX integration still needs one effective read/publish adapter
  and logical cache keys. Record 0026 documents those requirements and evidence.

## 0027 — DOCX patch predicate and Debug policy

- Primary main-document `Patch::is_noop()` replaces `changed()` without an alias;
  commit `Diagnostics::changed()` remains. Publishers and harness consumers are
  migrated, and both DOCX glossary profiles enforce the new name.
- Main-document snapshots, patches, operations, transfers, replacement inputs
  and local error/composition wrappers use content-free Debug formatting.
- Regression tests cover compact/pretty output and byte-vector disclosure.
  Display and shared-core diagnostics remain separate open policy work.
- Record 0027 contains verification and the remaining canonical owner scope.

## 0028 — explicit effective graph materialization

- Retained OPC candidates can materialize their effective graph into an owning
  package, preserving edits, additions/removals, relationships and content types.
- Conversion shares retained payloads, checks base/transfer fences and refuses
  managed base or donor escape before payload reads. Signature and ingress
  repair policy survives conversion.
- This supplies DOCX's future effective `to_owned_package` route. Canonical
  source ownership and logical caches remain open; record 0028 also records
  the empty-topology/read-ahead migration hazard.

## 0029 — source DOCX effective-owner boundary

- Same-lineage retained topologies can be adopted into immutable DOCX owners.
  Main-document views, explicit owning conversion and publication use the
  effective graph; logical owners receive fresh semantic caches.
- Pinned effective views retain payload admissions and donor execution guards.
  Queries observe cancellation while captured bytes survive source changes.
- All unmigrated physical operations refuse effective owners, including no-op
  requests. The raw package accessor is fallible; the unified consumer is adapted.
- Canonical source semantic Edit/Commit and specialized effective routes remain
  open. Record 0029 documents scope, verification and current allocation limits.

## 0030 — immutable retained semantic apply

- Source DOCX `Package::apply` returns a new effective owner without ZIP
  serialization/reopening. Snapshots and raw edits read effective main XML.
- Exact source proofs and effective settings policy guard changes; unrelated
  graph edits survive, no-ops preserve signatures, and inverses restore bytes.
- OPC full-document replacement/provenance primitives retain managed admission
  and source fences. Revision direct/replay edits preserve unmanaged identity.
- Canonical package-owning source Edit/Commit, specialized effective routes and
  allocation/performance work remain open. Record 0030 documents verification
  and the exact-byte (rather than strict logical-version) precondition boundary.


## 0031 — canonical source package-owning commit

- Source `Package::edit()` returns `document::Edit<source_backed::Package>`;
  commit returns the new effective owner through the canonical package accessors.
- Generic Edit/Commit share eager/source staging and diagnostics with private
  construction and specialized commit policy. Unified editor consumers migrate.
- Legacy stream/history harnesses explicitly request raw XML transactions.
- Canonical MCE coverage exposed and fixed escaped-namespace compatibility:
  unmanaged full-document apply uses the existing publication audit, while
  managed edits keep typed splice authority and admissions.
- Earlier pending primary-source-triad notes are superseded by this record.
  Specialized families, history, limits, performance and full acceptance remain
  open. See 0031 for scope and verification.


## 0032 — typed primary transaction limits

- `document::LimitKind` and `Error::Limit { kind, observed, maximum }` replace
  free-form resource strings and legacy count field names across primary paths.
- Durable aggregate transfer bytes now fail with a typed limit before payload
  copying; dynamic count/path helpers carry typed dimensions.
- DOCX glossary profiles ban the retired fields at their precise variant scope.
  Public boundary/rollback and durable overflow tests accompany the migration.
- Caller-selected semantic limits, other error families, policy vocabulary and
  full family registration remain open. See 0032 for verification and scope.


## 0033 — metadata on retained source owners

- Source DOCX metadata now follows effective property payloads and relationships,
  including canonical commit owners, additions/redirection and removal.
- A common OOXML prepared-graph adapter reuses existing property validation,
  repairs and decoding. Only the selected payload is decoded.
- Base/donor fences cover success and errors; repair reports stay per-owner and
  deduplicate repeated reads. Managed donor cancellation/release is tested.
- Other specialized reads and metadata writes remain pending. Record 0033
  documents verification, allocation limits and the remaining full-goal scope.


## 0034 — text streaming on retained source owners

- Source DOCX text output follows effective main replacements and relationship
  redirection, including canonical commit owners and managed donor payloads.
- Payload handles stay retained through output and final source checks; declared
  size is checked before capture. Partial-output errors keep exact progress.
- Canonical managed/unmanaged commits, donor cancellation, early limits and MCE
  source-failure precedence have integration coverage.
- Specialized reads/transactions, configurable limits and the wider goal remain
  open. Record 0034 documents verification and allocation boundaries.


## 0035 — shared OPC/OOXML SourceLimits vocabulary

- Renamed the physical profile and builder to SourceLimits/SourceLimitsBuilder
  throughout OPC, OOXML re-exports, unified consumers and active tools, without
  retaining aliases. Defaults, admission and checked builders are unchanged.
- Added minimal/full OPC, PPTX, XLSX and XLSB glossary profiles. Exact-scope bans
  distinguish XLSB's root re-export from unrelated semantic ReadLimits types.
- Broader tests caught an empty-plan publication regression: the one-shot
  path now preserves its exact-copy budget without allocating a retained owner.
- Kind/error names, options envelopes and other limit migrations remain open.
  Record 0035 documents checks and the unrelated existing iWork example error
  encountered by the exploratory workspace all-targets build.


## 0036 — typed physical source-limit errors

- OPC and OOXML expose SourceLimitKind; primary OPC failures use Limit with
  kind/observed/maximum and InvalidLimit with kind/value/maximum.
- Builder failures carry scalar or cross-field bounds. The members/entries
  inconsistency now identifies the excessive member dimension correctly.
- ZIP mapping, cached failures, unified classification and consumer assertions
  retain typed numeric evidence. Glossary profiles ban the retired vocabulary.
- Remaining configuration/dimension and semantic-limit work is documented in
  record 0036; the wider goal remains active.


## 0037 — canonical physical source dimensions

- Shared OPC/OOXML limits use members, entry_bytes, aggregate_bytes, depth and
  events with matching max_* methods and SourceLimitKind variants; no aliases.
- Registered all 24 OPC dimensions and their distinct meanings in the glossary.
  Defaults, admission rules and owner-local semantic profiles are unchanged.
- Migrated consumers, tests, active fuzz/performance tools and added exact-scope
  glossary bans. Record 0037 documents verification and remaining work.
- OpenOptions, semantic profiles and the wider unified goal remain active.


## 0038 — owning DOCX OpenOptions

- Added one crate-root OpenOptions envelope for owning DOCX ingress, canonical
  byte-slice constructors and configured password routes; removed the replaced
  limit-taking and ad-hoc constructors without aliases.
- Migrated unified decrypted adoption and consumers. The same source policy
  reaches clear OPC parsing while encryption mode remains retained.
- Added path/reader/bytes boundary and independent encryption/source limit
  checks; glossary rules ban the retired owning-Package methods.
- Record 0038 records evidence and remaining lazy-owner, semantic and crypto
  policy work. The broader unified goal remains active.


## 0039 — DOCX positional options and resource retention

- Replaced 17 path/positional constructor variants with canonical open and
  from_read_at options routes, using the existing root OpenOptions envelope.
- Cache, read-ahead and execution settings reach OPC admission and deferred
  reads; cloned options share budgets and cancellation. Path cancellation is
  checked before file opening.
- Owning routes explicitly refuse lazy policies before I/O, with typed unified
  classification. Sequential reader accounting and owner consolidation remain
  required; this does not declare the ingress family complete.
- Record 0039 documents migrated consumers, verification and remaining work.


## 0040 — managed sequential DOCX ingress

- Lazy reader and byte-slice ingress use OpenOptions; the old limits-only
  reader constructor is deleted and active consumers are migrated.
- Retained source buffers own their memory reservations through package clones
  and artifacts. Growth admits old/new peaks; reads, retries, copies and input
  counts obey caller execution policies, including the bounded EOF probe.
- Added focused failure/lifetime tests, a managed sequential fuzz target and a
  deterministic 1,024-mode smoke driver. Record 0040 documents results.
- Owner consolidation, semantic/crypto policy migration and the wider unified
  goal remain active; no complete milestone or performance gain is claimed.


## 0041 — lazy DOCX canonical publication and atomic save

- Replaced lazy Package::to_stream with write_to and migrated callers; added
  save (Full durability) and save_with_durability over retained effective graphs.
- Shared OPC atomic staging admits its fixed buffer to the destination owner's
  context and checks caller source state after file sync, before replacement.
  Pre-rename failures preserve the destination; post-rename sync errors retain
  their Committed classification.
- Record 0041 documents durability, preservation, refusal and cleanup tests.
- Owning/specialized output names, budgeted to_bytes, owner consolidation and
  the wider unified goal remain active. No complete family is claimed.
