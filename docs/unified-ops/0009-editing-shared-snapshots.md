# 0009 — Editing a shared snapshot (R4): share the base, never move it

- Status: decision record (analysis and recommendation; provisional until the
  owner confirms). No production code changed. One temporary probe test was
  built, run and deleted (§2.4).
- Date: 2026-10-05
- Scope: risk R4 of record 0003 (§2.2, §6) and the "Cross-cutting M2
  prerequisite (R4)" section of `m0a-log.md`. M2's write verbs depend on it.
- Base: `uo/0009-shared-edit` at `37ba0cbf3a` (= `uo/m0a` before S5d). Every
  `path:line` was read in that tree, except the OLE handles and the S5d leaf
  constructors, which were read at `2bec864ec1` (S5d, merged into `uo/m0a` as
  `cdc0f44c30`). *Inferred* marks a conclusion read from code but not executed.

Path prefixes: `O/` = `crates/litchi-opc/src/`, `D/` = `crates/litchi-docx/src/`,
`P/` = `crates/litchi-pptx/src/`, `X/` = `crates/litchi-xlsx/src/`,
`B/` = `crates/litchi-xlsb/src/`, `S/` = `crates/litchi-ops/src/`,
`perf/` = `docs/performance/`.

## 0. Decisions (provisional, listed for owner confirmation)

No guard in any leaf needs a unique owner. The OOXML publishers take `self` by
API choice, and their bodies only borrow. The packages are not `Clone` because
their index, catalog and cache are owned inline, not because of a lineage
token or exact-source authorization: one opened package has one lineage, which
every share carries by pointer. The OLE leaves are already shareable.

| # | Decision | Section |
|---|---|---|
| D1 | Every unified write verb takes `&self` and returns a new handle. No verb consumes a handle, calls `Arc::try_unwrap` or branches on a reference count, so the same operations on the same snapshot always publish the same bytes. | §1, §7 |
| D2 | Option A, "share the base, never move it". S0: `litchi_opc::SourceBackedPackage` becomes O(1) `Clone` (index, catalog and cache shared through `Arc`; one lineage per open). Consuming publishers keep their signatures and consume a share. | §2.1, §4, §6 |
| D3 | Leaves follow in tier order: xlsx needs no leaf change; docx publishers take `&self`; pptx builds its editor from `&self`; eager owners (docx, pptx, xlsb) apply into a new value instead of `&mut self`; the OLE leaves need nothing. | §2, §6 |
| D4 | Retire the xlsx shared fallback (materialize through `to_opc_package`) before M2 ships writes. On every real-producer fixture probed it broke the exact no-op (0 of 10–13 members kept), and it made output depend on sharing. Promote through the exact source instead. | §2.4 |
| D5 | Reject B (re-open from a retained reader) and C (consuming verbs with a sole-owner fast path). B fails the single-op budget on every shared write (pptx re-open 2,835 µs ≈ 42 % of a whole 6.7 ms one-edit cycle), and C's fast path never fires from Python. | §4 |
| D6 | No OOXML source-backed commit chains before publication. Until a leaf supports overlay chaining (option D, the follow-up), M2 rebases once (publish the pending commit into memory, re-open those bytes) at a family boundary or a second write, never per item. | §2, §5 |
| D7 | The batch executor works on one share per batch (zero conversion cost), builds one editor per run of same-family items, and meets the 2× budget by picking an engine that takes the whole run in one commit. | §5 |

---

## 1. The problem, stated precisely

A unified handle is one `Arc` over a leaf owner (record 0003 §2.2;
`S/format/docx.rs:56`, `S/format/pptx.rs:59`, `S/format/xlsx.rs:84`). Every
write returns a new handle and the old snapshot stays readable (design §3, §8;
ADR 0033 §3, `docs/adr/0033-unified-operations-api.md:63-65`). ADR 0003 asks for
"immutable, cheap-to-share `Send + Sync` snapshots" and allows mutation that
"consumes or borrows a snapshot into an `Edit`"
(`docs/adr/0003-snapshots-edits-and-patches.md:8,13`). The CRUD gate requires
"Cloneable snapshots share immutable source state"
(`docs/CRUD_Scenario_Checklist.md:30`).

A write needs the leaf owner at up to three points:

1. **Edit start**: building the leaf editor. docx and every eager owner start
   from `&self`; the pptx source owner and the xlsx source owner move their
   package into the editor.
2. **Apply**: eager owners fold a commit into the package with `&mut self`
   (docx, pptx, xlsb).
3. **Publication**: every OOXML source-backed publisher takes `self`, down to
   `litchi_opc::SourceBackedPackage`.

Where a point needs ownership, a shared handle cannot provide it. Sharing is
the common case, not the exception:

- **Python.** A PyO3 method receives the handle object by reference
  (`PyRef<Self>`). To hand the inner Rust handle to a consuming verb it must
  either clone it (two references) or take it out of the Python object, which
  would break the old object the design says stays usable (design §8). And the old
  object is still alive during `doc = doc.set(...)`: the rebinding happens after
  the call returns. A sole-owner fast path therefore never fires from Python.
- **Rust.** `let f2 = f.set(..)?` keeps `f`, so a consuming verb needs
  `f.clone()` first. Lazy readers hold the inner state too: every `SourceSlide`
  (`P/presentation/source.rs:431-434`) and `SourceWorksheet`
  (`X/workbook/source.rs:123-126`), and so every `SelectedWorksheet` in
  litchi-ops (`S/format/xlsx.rs:403-406`), keeps the count above one.

Today's M0a handles therefore behave as follows when shared: xlsx falls back
to an eager snapshot built from the same index (`S/format/xlsx.rs:418-443`);
pptx refuses with no error (`S/format/pptx.rs:505-552`); docx can start an edit
but cannot publish it (`S/format/docx.rs:388-404`; `D/source_backed.rs:1500`).

## 2. Findings per leaf

### 2.1 `litchi-opc` and the positional source (the root of the OOXML leaves)

**What the package holds.** `SourceBackedPackage` (`O/source_backed.rs:6453-6476`)
owns `source: SourceSnapshot`, `archive: IndexedArchive<SourceReader>`, an
optional `Arc<ArchiveReadAhead>`, limits, the content-types member name, the
package relationships, the catalog (`parts: Vec<CatalogPart>`, relationship
metrics, `parts_by_name`, `casefold_order`, `non_part_members`), the tolerance
report, `cache: PartCache` and an optional `Arc<Reservation>`.

- Already shared: `SourceSnapshot` is `Clone` and holds `Arc<dyn ReadAt>`, the
  version, an `Arc` monitored-read depth and the lineage
  (`O/source_backed.rs:3691-3701`); `SourceReader` is `Clone`
  (`:3586-3594`); the read-ahead state is an `Arc` (`:6456`).
- Owned inline, so not `Clone`: the archive index (soapberry-zip
  `office.rs:1703-1734`, `ZipArchive` + `Vec`/`HashMap` + `OnceLock`), the
  catalog vectors and maps, and `PartCache` (`O/source_backed.rs:5317-5338`,
  `#[derive(Debug)]` only, `Mutex` state).
- Built for concurrent `&self` use: the cache state is behind a `Mutex`
  (`:5320`); monitored-read scopes "count rather than latch so that concurrent
  operations on the same snapshot, a nested scope, and a scope taken on a clone
  of the snapshot cannot end one another's monitoring" (`:3635-3638`); the
  managed write session and its permits are `Mutex`-guarded (`:5328-5330`,
  `:5519-5562`). Reads of one package from several threads are tested
  (`std::thread::scope` over `&package`, e.g. `:18494`).

**What publication consumes.** Ten public doors take `self`:
`write_topology_to_stream` (`:9905`), `write_part_overlay_to_stream` and its
shared/accounting/relationship-removal variants (`:9979`, `:9997`, `:10023`,
`:10133`), and the bounded multi-part doors (`:10243`, `:10486`, `:10505`,
`:10537`, `:10552`); also `prepare_topology(self)` (`:8022`) and
`into_opc_package(self)` (`:7853`). Their bodies only borrow:

- `write_exact_source(self)` reads `&self.source` and the cache context
  (`:11675-11677`);
- the overlay writer borrows the archive
  (`self.archive.preservation_index_with_limits(..)`, `:11770-11772`), clones
  the snapshot into its sink (`SourceCheckedSink`, `:12007`, `:12061`) and
  reads the cache (`:11985`);
- `publish_prepared_topology` takes the package by value and only borrows it
  again (`:2831-2888`);
- the one field move is `into_opc_package`'s `mem::take` of
  `non_part_members` (`:7884-7885`); the borrowed `to_opc_package(&self)`
  (`:7907`) copies them instead.

**Borrowing doors that already exist.** `prepare_source_part_splice(&self, ..)`
returns a plan that borrows the package; the plan, not the package, is consumed
by `write_to_stream` (`O/source_backed/splice.rs:1101-1110`, `:491`; used by
docx tail append, `D/source_backed/tail_append.rs:690`, `:756`).
`with_prepared_topology(&self, ..)` validates a candidate and leaves the
package usable (`O/source_backed.rs:8030-8061`).
`source_artifact().write_to_stream(&self)` copies the exact source
(`:7640`, `:3829`).

**Why `self` was chosen.** The only stated rationale is change 0613
(`perf/0613-opc-original-audit-memo.md:264-271`): consuming `self` "is what
makes 'the source this package was opened on has not been published over'
structural rather than checked, and it is what lets publication drop
read-ahead, retire reservations and hand the archive to the writer.
Re-entrant publication needs its own record." This record is that record.
Against the code at this base:

- *Not published over.* Every read and every publication sink write already
  re-checks the source version (`SourceSnapshot::ensure_current`,
  `:3855-3865`; monitored publication, `:3885`; `SourceCheckedSink`,
  `:12007`). A path source pins its open descriptor: "The open descriptor
  remains pinned to its original inode across an atomic pathname replacement"
  (`crates/litchi-core/src/source/file.rs:25-28`). So an atomic save over the
  source path leaves every handle reading the old, unchanged inode, and an
  in-place overwrite turns the next read into a typed `SourceChanged`. The
  property is already *checked*; consumption adds nothing to it. On Windows the
  source is reopened with `FILE_SHARE_READ | FILE_SHARE_WRITE |
  FILE_SHARE_DELETE` (`file.rs:283`) and versioned by timestamps; replacement
  over an open handle is not exercised by any test (not claimed).
- *Drop read-ahead.* `disable_for_exact_publication(&self, ..)` is a
  "monotonic" switch of "the shared archive adapter"
  (`O/source_backed/read_ahead.rs:417-438`); a plain `&self` read sets it too
  (`content_types_data_with_limit`, `O/source_backed.rs:7228-7229`). It is a
  speed effect.
- *Retire reservations.* Dropping the package ends its reservations. With
  sharing, they end when the last handle drops. Handles already keep the
  package alive for reads, so nothing new is retained (§4, memory).
- *Hand the archive to the writer.* The writer borrows it (`:11770`).

**Identity guards.** `SourceLineage(Arc<()>)` is compared by pointer and "lets
a semantic facade bind snapshots and patches to this opened package rather
than to a merely equal `SourceVersion`" (`:3703-3717`, `:7614-7622`). Every
open mints a fresh lineage (`:6515`, `:6871`). Handles that share one package
share its lineage, so every lineage guard passes; a re-open (option B) gets a
new lineage and every outstanding commit becomes stale. One check compares
package *instances*: `prepare_source_part_splice_with_fragment` refuses when
`!std::ptr::eq(self, fragment.package)`
(`O/source_backed/splice.rs:1079-1083`). It holds for one package shared
through an `Arc`, and refuses across two separate clones.

**Exact-source authorization.** ADR 0005's 2026-08-21 amendment makes the
retained exact source plus an unrevoked authorization the only authority for
byte-identical passthrough; "any mutable OPC seam revokes that authorization"
(`docs/adr/0005-io-memory-and-performance.md:344-353`).
`SourceBackedPackage` has no mutable seam; edits travel as separate plans and
an empty plan copies the source (`O/source_backed.rs:10581-10582`). Sharing
adds no seam. The eager `OpcPackage` is the precedent for sharing with
copy-on-write publish state: it is `#[derive(Clone)]`
(`O/package.rs:382-383`), its authorization is "clone-local"
(`:424-425`), and a clone shares the owned source while revoking
independently (test `clone_shares_owned_source_but_revocation_is_independent`,
`:3854-3889`; ADR 0030 "Clone, Send and Sync").

**Materialization loses the exact source.** `to_opc_package(&self)` builds a
fresh `OpcPackage::new()` and adds every part through `try_add_source_part`
(`O/source_backed.rs:7957-7993`), which revokes exact-source authorization
(`O/package.rs:1919-1920`) and never calls `authorize_owned_source`
(`:2625-2640`). The result has no retained archive and no preservation
provenance, so its save regenerates every ZIP member. §2.4 measures this.

**Verdict.** The package is not `Clone` for incidental reasons (owned index,
catalog and cache). Its publishers take `self` by API choice; the bodies
borrow. No guard needs a unique owner. The smallest fix is an O(1) `Clone`
that shares the index, catalog and cache, so a publisher consumes a share of
the one opened package (§4, A-clone); `&self` publishers are the equivalent
fallback.

### 2.2 DOCX

**Source-backed owner.** `source_backed::Package` (`D/source_backed.rs:124-136`)
holds the `SourceBackedPackage` by value, `Option<ExecutionContext>`,
`Arc<ParagraphIndexCache>`, `Mutex<ToleranceReport>` and an `AtomicBool`. It is
not `Clone`; the package is the hard blocker, and the two interior-mutable
fields would need a hand-written clone that copies them together (inferred).

- *Edit start borrows.* `edit_document(&self)` (`:1142`) captures the main
  document through `main_document_snapshot_with_before` (`:1962`);
  `Snapshot::edit(&self)` clones the snapshot (`D/document/transaction.rs:1476`).
  One `Edit` holds many operations (`&mut self` operations, e.g. `:2273`).
- *Commit.* `Edit::commit(self)` (`:5059`) returns `Commit { snapshot, patch,
  diagnostics }` (`:5265-5270`). The snapshot's XML storage carries
  `SourceIdentity { lineage, version, partname }` (built at
  `D/source_backed.rs:2077`); it never references the docx `Package`, the
  catalog or the cache.
- *Publication consumes.* All 32 source-backed publishers take `self`; 23 move
  `self.package` into a consuming OPC door, and the 9 inverse publishers only
  call `&self` methods (`source_artifact().write_to_stream`,
  `restore_source_artifact_to_stream`). The main one,
  `publish_document_commit_to_stream` (`:1500-1556`), re-reads the current main
  part, applies the patch, validates, refuses transfers (`:1512-1524`), and
  only then moves the package into one of three OPC doors: an empty overlay set
  for an exact no-op (`:1534-1537`), `write_topology_to_stream` for a
  source-XML proof (`:1538-1544`), or `write_part_overlay_shared_to_stream`
  (`:1546-1552`). Tail append already publishes from a borrowed package
  (`Edit/Plan/Commit<'package>` over `&Package`,
  `D/source_backed/tail_append.rs:340-441`).
- *No chaining.* A second commit started from `commit.snapshot().edit()` cannot
  publish: the publisher re-reads the package's main part and `Patch::apply`
  requires the patch's `before` to match it, so it returns `StaleSource`
  (`D/source_backed.rs:1506-1511`, `D/document/transaction.rs:5357-5360`,
  `:1494-1507`; inferred). Many operations fit in one `Edit`; several commits
  need publish → reopen → edit.
- *Families.* 14 source-backed families, all edit from `&self` and publish by
  `self` except the two tail-append families (main document, variables,
  section layout, story text, glossary batch, paragraph copy and removal, SVG,
  hyperlink sanitize/redact, story-hyperlink redact, altChunk, tail append and
  streamed tail append).
- *What the docx records say.* Change 0039 introduced "the consuming
  `publish_document_commit_to_stream`" with no reason beyond the OPC door it
  calls (`perf/changes/0039-docx-source-backed-semantic-publication.md:19-21`);
  the only code comment on consumption is the redaction publisher's "This
  operation consumes the package and exposes no inverse API"
  (`D/source_backed.rs:1717-1718`). Change 0497's consuming `write_to_path`
  consumes a plan or commit, not the package
  (`perf/changes/0497-docx-atomic-publication.md:20-33`).

**Eager owner.** `litchi_docx::Package` (`D/package/model.rs:133-156`) wraps a
`Clone` `OpcPackage` but is not `Clone` itself (`properties: Slot`,
custom properties and `Option<MutableDocument>` are not `Clone`).
`edit_document(&self)` borrows (`D/package/package/document.rs:36-40`).
Applying needs `&mut self` (`apply_document_patch` `:99`,
`publish_document_commit` `:168`, and the rest); they all go through
`edit_opc_candidate` (`D/package/package/access.rs:218-261`), which clones the
OPC package, edits the copy and assigns it back. So the fork already exists;
it only needs to return a new `Package` instead of assigning. `save` and
`to_stream` take `&mut self` (`D/package/codec.rs:522`, `:606`) but mutate only
for a mutable document or dirty properties (`:1281`, `:1299`); the final write
is `&self` (`O/package.rs:2423`).

**Identity guards.** Transaction identity (`D/source_backed.rs:1982`;
`D/document/transaction.rs:865-896`), section layout (`D/section/layout.rs:543-570`),
story text, SVG and story hyperlinks (lineage + version + artifact fingerprint,
`D/story_hyperlinks.rs:854-877`) all compare lineage by pointer.
One shared package passes all of them. Per-patch one-shot gates exist only in
eager families (tracked revisions, conflicts, content controls) and gate the
patch, not the package.

**Not R4 but on the same path.** Source-backed `edit_document()` refuses
Word-authored files that declare `mc:Ignorable` (m0a-log S5a): the unmanaged
route raises `UnsafeEdit` at `D/source_backed.rs:2068-2073` when MCE
processing rewrites the main part; the managed route refuses any `mc` binding
(`:2861-2908`). Record 0008 owns that fix.

### 2.3 PPTX

**Source-backed presentation.** `SourceBackedPresentation` is already a cheap
clone: `#[derive(Clone)]` over `Arc<SourceInner>` (`P/presentation/source.rs:412-415`).
`SourceInner` (`:112-129`) holds the `SourceBackedPackage` by value, the
presentation part, `Box<[Arc<SourceSlideData>]>`, limits, the tolerance report
and a flag. Each `SourceSlide` also holds the `Arc<SourceInner>` (`:431-434`).
The reader is retained only inside the OPC snapshot; nothing public re-opens
from it, and `PptxState` keeps no reader (`S/format/pptx.rs:67-77`).

- *Edit start consumes.* `try_into_editor(self)` (`:1216-1257`) does all
  fallible work borrowed (`editor_seed(&self)`, `:1261-1289`), then needs
  `Arc::try_unwrap(self.inner)` (`:1228`) to move the package into the editor.
  Its documentation states the reason: "The editor takes ownership of the
  package because publication consumes it, so the conversion succeeds only
  while this value is the sole owner of the opened source: no clone of it and
  no `SourceSlide` from it is alive" (`:1205-1208`). litchi-ops adds its own
  `Arc::try_unwrap` (`S/format/pptx.rs:521`).
- *Editor.* `SourceBackedPresentationEditor` (`:687-698`) holds the package by
  value plus `Arc`-shared catalog data. It is "intentionally not cloneable:
  publishing consumes its deferred OPC source to ensure that the exact source
  checked during editing is the source raw-copied to output" (`:679-686`).
  That guarantee rests on lineage and version identity, which every share of
  one package has by construction; it does not need a unique owner.
- *Publication.* Nine publishers take `self`: `publish_slide_commit_to_stream`
  (`:1920`), SVG (`:1967`, alias `:1993`), slide batch (`:2006`), slide order
  (`P/presentation/order.rs:144`, `:158`), cross-slide copy
  (`P/presentation/source_cross_copy.rs:445`) and SVG attachment
  (`P/presentation/source/svg_lifecycle.rs:504`, alias `:531`). Every body
  borrows until a final `self.package.write_*` call (e.g. `:1945`).
- *No chaining.* Each edit holds one operation (`:2484-2489`). A commit
  (`SourceBackedSlideCommit { snapshot, patch }`, `:729-732`, `:2580`) is bound
  to its source by `same_source` (position, part, version, lineage, bytes and
  closure; `:2439-2450`), and publication re-reads the current slide from the
  package (`:1926-1941`), so a chained commit is `StaleSource`. The batch is
  at most 32 slides, each once (`MAX_SOURCE_BACKED_SLIDE_BATCH`, `:100`;
  `:2928-2932`). Six transaction kinds:
  single slide, SVG replacement, multi-slide text batch, slide order,
  cross-presentation copy and SVG attach/detach; no notes on the source path.
- *Cross-copy self-reference.* Copying between two handles of one
  presentation shares lineage and version, which cross-copy refuses
  (`P/presentation/source_cross_copy.rs:780-784`). Sharing makes the case
  reachable; it still fails closed.

**Eager owner.** `litchi_pptx::Package` (`P/package/model.rs:13-48`): every
field is `Clone` (`OpcPackage`, `Option<MutablePresentation>`,
`ToleranceReport`, `OnceLock<Arc<PartDigests>>`, encryption state), but the
type does not implement `Clone`, and a plain derive would be wrong: the digest
memo must be re-projected onto the clone's own allocations (`memo_for`,
`P/package/model.rs:1897-1912`; ADR 0005 2026-09-16 amendment and ADR 0032).
`opened_presentation_transaction(&self)` borrows (`:296`); the transaction
clones the `OpcPackage` (`P/opened/transaction.rs:42-50`) and its commit
snapshot holds an `Arc<OpcPackage>` (`P/opened/model.rs:298-337`). Applying
(`apply_opened_presentation_commit`, `:550`) and saving
(`save`/`to_bytes`, `P/package/codec.rs:319`, `:360`) take `&mut self`.

**Memos.** ADR 0005 (2026-09-16) and ADR 0032 bind a memo to the allocations
the owning snapshot holds and say a facade adopts the memo of each published
snapshot. Immutable handles comply: each keeps the memo of its own graph, and
handles of one snapshot share it (ADR 0032: "a clone shares it",
`docs/adr/0032-snapshot-derived-value-memos.md:130-131`).

### 2.4 XLSX (with a probe of today's shared fallback)

**Source-backed workbook.** `SourceBackedWorkbook` is already a cheap clone:
`#[derive(Clone)]` over `Arc<SourceInner>` (`X/workbook/source.rs:116-119`).
`SourceInner` (`:77-105`) holds the `SourceBackedPackage` by value plus
`OnceLock` caches and `Box<[Arc<SourceSheetData>]>`. Every `SourceWorksheet`
holds the same `Arc<SourceInner>` (`:123-126`). The L-XLSX-1 accessor
`try_into_source_backed_package(self)` is `Arc::try_unwrap` and "returns `self`
unchanged while another clone of this workbook, or a `SourceWorksheet` handle
selected from it, still shares the catalog" (`:645-654`).

**Per-feature source editors.** Sixteen editors (cell values, merges, defined
names, hyperlinks, tab state, row visibility, auto filter, calculation
properties, conditional formatting, data validation, page breaks, margins,
setup, print options, sheet protection, form controls) each own the package by
value and are built by `from_source_backed_package(package)` (e.g.
`X/cell_values/source.rs:442-445`). Edits start from `&self`
(`edit`, `edit_sheets`, `edit_many`, `:454`, `:470`, `:479`); every publisher
takes `self` (`publish_commit_to_stream`, `:537`; `publish_multi_commit_to_stream`,
`:600`; one per other editor) and borrows until its consuming OPC call.

- *No chaining.* `edit()` always starts from the source, and publication
  requires the commit's `before` to equal the current source
  (`X/cell_values/source.rs:553-572`; `X/defined_names/source.rs:147-153`).
  Change 0635: "No public API seeds a value-only edit from a post-commit
  snapshot"; the only chain is publish then reopen
  (`perf/0635-xlsx-facts-builder-and-chains.md:340-356`).
- *No cross-family publication.* Cell values and defined names both rewrite
  `xl/workbook.xml` (`X/cell_values/mod.rs:7-9`, `X/defined_names/source.rs:136-141`),
  and OPC refuses two overlays on one member (`O/source_backed.rs:11816-11820`).
- *Caps per commit.* 256 unique cells and expanded actions
  (`X/cell_values/source.rs:24`), 64 worksheet owners
  (`X/cell_values/snapshot.rs:139`), single-worksheet workbooks only for
  value-only edits (`:562`), and no worksheet with a table relationship
  (`:2059`); row visibility and hyperlinks also cap at 256.
- *Editor-bound commits.* The tab-state editor stamps `origin: Arc<()>` into
  its commits, so only the editor that produced a commit can publish it
  (`X/tab_state/source.rs:121-128`, `:158-163`). A handle that defers
  publication must keep that editor with the pending commit.
- *Managed handoff.* Change 0151 makes a managed payload that is asked to
  escape into an owning `Arc` fail with `ManagedPartDataArcEscape` instead of
  detaching its reservation
  (`perf/changes/0151-xlsx-managed-source-editors.md:24-40`). That protects
  budget integrity; it says nothing about how many handles share the package,
  and a share of the package keeps every reservation attached.

**Eager workbook.** `Workbook` is the immutable snapshot ADR 0003 describes:
"Immutable, cheap-to-share XLSX workbook snapshot", `#[derive(Debug, Clone)]`
over `Arc<Inner>` (`X/workbook/model.rs:421-427`). `edit(&self)` (`:1313`)
starts a multi-feature `Edit`; `commit(self)` (`X/workbook/edit/semantic/transaction.rs:1501`)
returns a new workbook, commits chain, and `Edit::join` merges transactions
(`:1279`). Saving takes `&self` (`X/workbook/model.rs:1156-1191`). Its output
is byte-preserving only when the package retains its owned source: then an
exact no-op writes the source bytes (`O/pkgwriter.rs:1258-1265`) and untouched
members are raw-copied; otherwise every member is regenerated. Signed workbooks
are refused at `Edit::new` (`transaction.rs:336-337`).

**Today's shared fallback** (`S/format/xlsx.rs:418-443`, `:459-461`):
`Package::from_opc(package.to_opc_package()?)?.into_workbook()?`.

- It fails for every managed handle: `to_opc_package` refuses with
  `ManagedPackageMaterialization` (`O/source_backed.rs:7926-7935`).
- It refuses signed workbooks (eager `Edit::new`), where the source path still
  publishes a signed exact no-op.
- It decodes every part. The matched eager control of change 0046 (positional
  open, complete materialization, owning transaction, ordinary writer) took
  215.5 ms against 1.61 ms on the source path (133.67×), and change 0061's
  216.8 ms against 4.65 ms (46.65×), both on the 16.8 MB media fixture
  (`perf/changes/0046-xlsx-source-backed-calculation-metadata-publication.md:73-76`,
  `perf/changes/0061-xlsx-source-backed-page-break-publication.md:67-70`).
- It loses the exact source, so its output is not byte-preserving (§2.1). The
  probe below measures that.

**Probe.** A temporary test in `litchi-xlsx` (built at this base, debug
profile, then deleted) opened each fixture from `OwnedSource` bytes and
published it two ways: the *sole-owner* route
(`try_into_source_backed_package`, then an empty overlay set, or a
`cell_values::SourceBackedEditor` setting A1 to 2) and the *shared* route
(the `materialize` above, then `to_bytes`, or an eager `Edit` setting A1 to 2).
"Raw members" counts ZIP members whose local header and data span is
byte-identical to the source's (the `raw_members` helper of
`crates/litchi-xlsx/tests/source_backed_hyperlinks.rs:116-145`).

| Fixture (`test-data/ooxml/xlsx/`) | Operation | Sole-owner route | Shared route |
|---|---|---|---|
| `sheet-names.xlsx` (9,425 B) | exact no-op | identical, 13/13 raw | 7,516 B, 0/13 raw |
| `hyperlinks.xlsx` (9,068 B) | exact no-op | identical, 13/13 raw | 7,251 B, 0/13 raw |
| `InlineString.xlsx` (8,401 B) | exact no-op | identical, 10/10 raw | 6,424 B, 0/10 raw |
| `InlineString.xlsx` | A1 := 2 | 8,376 B, 8/10 raw | 6,380 B, 0/10 raw; differs from the sole-owner output |
| `table-sample.xlsx` (10,595 B) | exact no-op | identical, 13/13 raw | 8,568 B, 0/13 raw |
| created by `Workbook::create` (2,159 B) | exact no-op | identical, 6/6 raw | identical, 6/6 raw |
| created by `Workbook::create` | A1 := 2 | 5/6 raw | 5/6 raw; same bytes as the sole-owner output |

On every real-producer fixture the shared route breaks the exact no-op the CRUD
gate requires (`docs/CRUD_Scenario_Checklist.md:55`) and rewrites every
untouched member; the same edit produces different bytes depending on whether a
second handle exists. Only Litchi's own output, whose compression the eager
writer reproduces, hides the difference, which is why the benchmark corpora
(all Litchi-written) never showed it. The source value-only editor refused the
edit on three of the four real fixtures ("value-only edits currently require
exactly one worksheet", `X/cell_values/snapshot.rs:562`; a table relationship,
`:2059`), so most xlsx writes in M2 will run on the eager engine anyway.

**A fidelity-preserving promotion already exists in the public API** (inferred,
not run): `source_artifact().write_to_stream(&mut buffer)`
(`O/source_backed.rs:7640`, `:3829`) followed by `Workbook::from_bytes(buffer)`
(`X/workbook/model.rs:477`) gives an eager workbook that owns its exact source,
so exact no-ops stay exact and untouched members are raw-copied. It costs one
copy of the archive plus an in-memory open (`xlsx_open_owned` 104.8 µs on the
medium corpus, 1.348 ms on dense-wide; `perf/0744-xlsx-eager-workbook-cell-path.md:292`, `:298`).
For a managed handle it avoids the `ManagedPackageMaterialization` refusal,
because it copies the archive instead of materializing reserved payloads, but
the copy and the eager workbook would then sit outside the caller's budget; M2
must charge them to the handle's execution context or refuse (open point, §9).

### 2.5 XLSB

- *Source workbook.* `SourceBackedWorkbook` is `Clone` over `Arc<SourceInner>`
  (`B/workbook/source.rs:115-118`) and has no edit or publish API, and no
  package accessor; no OPC publisher is called anywhere in `litchi-xlsb`.
- *Materialization.* `to_workbook(&self)` is `to_opc_package` plus a complete
  eager parse (`B/workbook/source.rs:995-1006`); it refuses managed handles
  (`:989-993`) and yields a package without its exact source, so a later save
  regenerates every member, as in §2.4. litchi-ops calls it on every
  `Xlsb::editor()` call with no caching (`S/format/xlsb.rs:370-384`).
- *Eager workbook.* `xlsb::Workbook` (`B/workbook/model.rs:32`) is not
  `Clone`, although every field type is. Edits start from `&self`
  (`edit_cell_values`, `B/workbook/package.rs:380-382`); applying takes
  `&mut self` (`apply_cell_values`, `:418-433`), which already computes a new
  workbook through `apply_retaining_parse` and then overwrites `*self`. Save is
  `save(&self, W: Write + Seek)` (`:1033-1036`). `XlsbEditor::Eager(&Workbook)`
  in litchi-ops (`S/format/xlsb.rs:105-110`) therefore cannot apply a commit
  without a copy that the type does not offer.
- *Families.* About 15 per-family transactions; cells and workbook structure
  return `(Snapshot, Patch)`, others apply in place; several strip signatures
  on changed commits (record 0002, cross-cutting finding 4).

XLSB has no consuming source publisher, so its R4 problem is the eager apply:
it needs a non-mutating apply that returns the new workbook (the computation in
`apply_retaining_parse` already produces one), and a promotion that keeps the
exact source (copy the artifact, then open the eager package from those bytes)
instead of `to_workbook()`.

### 2.6 Legacy OLE: DOC, XLS, PPT, `litchi-cfb`, `litchi-ole-common`

R4 does not apply to the OLE leaves. Every edit owner is an immutable `Clone`
snapshot over `Arc`-shared bytes or an `Arc<SharedOleFile>`; every edit entry
takes `&self`; `commit(self)` consumes only the short-lived transaction; every
publication takes `&self` and can be repeated. Prefixes here: `doc/`, `xls/`,
`ppt/`, `cfb/`, `olec/` = `crates/litchi-{doc,xls,ppt,cfb,ole-common}/src/`.

| Leaf | Edit owner (shared state) | Edit start | Commit chains? | Publication |
|---|---|---|---|---|
| DOC eager | `body_text::Snapshot { source: Arc<[u8]>, .. }`, manual `Clone` (`doc/body_text.rs:609-633`) | `edit(&self)` (`:930`), but each start re-parses a full byte copy (`:1326-1333`) | yes (`Edit::commit` returns a new snapshot, `:2086`, `:2295`) | `bytes()`, `bytes_shared()`, `finish()`, all `&self`; no I/O sink |
| DOC source | `SourceSnapshot { inner: Arc<SourceInner> }`, derived `Clone` (`doc/body_text/source.rs:493-496`, `:441-452`) | `edit(&self, pos)` (`:890`) | yes (inferred: each commit nests one composed view, `:998`, `:1385`) | `write_to`, `publish_to_stream`, `save`, `save_with_durability`, all `&self` (`:1133-1150`) |
| XLS | `cell_values::Snapshot { inner: Arc<Inner> }` over `Arc<[u8]>` (`xls/cell_values/mod.rs:549`, `:535-546`) | `edit(&self)`, an `Arc` clone (`:795`) | `commit_source_backed` yes (`:2614`, target materialized at `:5464`); `commit_source_backed_plan` is forward-only (`:2646`, `:3482`) | `write_to`, `publish_to_stream`, `save`, all `&self` (`:3421-3447`, `:3535-3562`) |
| PPT | `slide_order::Snapshot` over `Arc<[u8]>` (`ppt/slide_order.rs:571`); `text_edit::SourceSnapshot { Arc<SourceInner> }` (`ppt/text_edit.rs:348-360`) | `edit(&self)` (`ppt/slide_order.rs:859`); `edit_text(&self)` (`ppt/text_edit.rs:478`) | slide order yes (`:2121`); source text yes (inferred) | `bytes(&self)`; source `write_to`, `save`, all `&self` (`ppt/text_edit.rs:711-728`) |

Substrate: `SharedOleFile` (`cfb/shared.rs:567-583`) holds `Arc<dyn ReadAt>`,
the version, an `Arc` parsed index and `Mutex`/atomic caches. It is used only
as an `Arc` and is `Send + Sync` (a test moves clones to other threads,
`cfb/shared.rs:4559-4572`). Overlay, splice and stream-move planners take
`&self` (`cfb/overlay.rs:776`, `cfb/splice.rs:179`, `:205`,
`cfb/stream_move.rs:237`), and a `ValidatedOverlayPlan` publishes from `&self`
(`cfb/overlay.rs:948-999`). `SourceBackedOverlayPublisher` holds
`Arc<SharedOleFile>` (`olec/source_backed_overlay.rs:22`) and plans from
`&self`. Identity guards are versions, lengths and SHA-256 fingerprints that
every share of one snapshot has in common, so sharing breaks none of them. The
only consuming entry, `xls` `chart_area::Snapshot::edit(self)`
(`xls/chart/chart_area/model.rs:36`), is on a `Copy` type.

What S5d already did (merged as `cdc0f44c30`): the `Doc`, `Xls` and `Ppt`
handles build their editors from `&self` over the retained
`Arc<SharedOleFile>`: `SourceSnapshot::from_shared(Arc::clone(shared))` for
DOC (L-DOC-1), `text_edit::SourceSnapshot::from_shared` and
`SourceBackedPackage::text_edit_snapshot` for PPT (L-PPT-1), and
`cell_values::Snapshot::from_shared_ole_file` for XLS (L-XLS-1, which copies the
source once and re-indexes it at the first edit, about 23 µs vs 21 µs on
`Simple.xls`, m0a-log S5d). `SourceBackedOverlayPublisher::from_shared` adopts
the shared index.

The OLE costs are elsewhere and do not depend on sharing: the DOC and PPT
source owners fingerprint the complete artifact at open and at commit
(`perf/changes/0105-doc-source-backed-paragraph-splice.md:60-66`;
`perf/CRUD_COVERAGE.md:1507`), and the XLS editors materialize one complete
target allocation (`xls/cell_values/mod.rs:3366-3371`). One requirement for
M2 follows: after the first write, the handle must hold the leaf *edit*
snapshot (which chains), not the read owner, so the XLS copy and re-index is
paid once per opened file, not once per write.

### 2.7 Per-leaf summary

| Leaf owner | Edit start | Apply / publication | `Clone` today | What blocks a shared write | Unique owner needed? |
|---|---|---|---|---|---|
| `litchi-opc` `SourceBackedPackage` | — | 10 public doors, `prepare_topology` and `into_opc_package` take `self`; bodies borrow | no: index, catalog and cache owned inline | receivers and owned fields | no |
| `litchi-opc` `OpcPackage` (eager) | — | `to_stream(&self)` | yes, O(parts), clone-local authorization | — | no |
| docx source `Package` | `edit_document(&self)` | 32 publishers take `self`; 23 move the OPC package | no (`SourceBackedPackage`, `Mutex`, `AtomicBool`) | OPC receivers | no |
| docx eager `Package` | `edit_document(&self)` | apply and save take `&mut self` (fork inside) | no (`Slot`, custom properties, `MutableDocument`) | `&mut` apply | no |
| pptx `SourceBackedPresentation` | `try_into_editor(self)` needs `Arc::try_unwrap` | 9 editor publishers take `self` | presentation yes (`Arc`); editor no, by intent | package moved into the editor | no |
| pptx eager `Package` | `opened_presentation_transaction(&self)` | apply, `save`, `to_bytes` take `&mut self` | no; memo must be re-projected | `&mut` apply and save | no |
| xlsx `SourceBackedWorkbook` + 16 editors | editors built from the package by value; `edit(&self)` | publishers take `self` | workbook yes (`Arc`); editors no | package moved into the editor | no (tab-state commits bound to their editor) |
| xlsx eager `Workbook` | `edit(&self)` | `to_bytes`, `save` take `&self` | yes (`Arc<Inner>`) | — | no |
| xlsb | eager `edit_cell_values(&self)`; no source editor | `apply_cell_values(&mut self)`; `save(&self)` | source yes; eager `Workbook` no | `&mut` apply; `to_workbook()` per call | no |
| doc, xls, ppt | `edit(&self)` on `Clone` snapshots | `&self`, repeatable | yes | — | no |
| `litchi-cfb`, `litchi-ole-common` | planners take `&self` | plans publish from `&self` | used through `Arc` | — | no |

## 3. Why the owners are not `Clone`: incidental, not a uniqueness requirement

Every candidate reason for a unique owner, checked against the code:

| Candidate reason | Where | Verdict |
|---|---|---|
| Lineage token | `SourceLineage(Arc<()>)`, pointer equality, minted per open (`O/source_backed.rs:3703-3717`, `:6515`, `:6871`) | Identifies the *opened package*, not an owner. Every share of one package has the same lineage; only a re-open needs a new one. |
| Exact-source authorization | ADR 0005 2026-08-21; eager `exact_source_authorized` is clone-local (`O/package.rs:424-425`, `:2572-2580`) | The source-backed package has no mutable seam that could revoke it. The eager package already authorizes per clone. |
| "The exact source checked during editing is the source raw-copied to output" | pptx editor docs (`P/presentation/source.rs:679-686`) | Enforced at publication by `same_source` (lineage, version, bytes; `:2439-2450`) and the version fences (`:1926-1941`). It holds for any share of one package. |
| "Not published over" | change 0613 (`perf/0613-opc-original-audit-memo.md:264-271`) | Already checked: version fences on every read and sink write, and a pinned descriptor across atomic replacement (§2.1). |
| One-shot publication tokens | `AuthorizedPrecompressedPart` (no public constructor, `O/source_backed.rs:211-226`); `RetainedPrecompressedPart`, "a shareable source-authorized compressed capture for repeated publication … cloning this handle does not authorize a writer" (`:278-282`) | Minted per publication from a shareable capture. This is the pattern the record generalizes: share the immutable capture, authorize each publication. |
| Per-patch one-shot gates | docx tracked revisions, conflicts, content controls (eager families; `D/revision/authoring.rs:757-815`) | Gate a patch, not the package. |
| Package-instance check | `std::ptr::eq(self, fragment.package)` (`O/source_backed/splice.rs:1079`) | Needs the fragment and the plan to come from one value, which they do in every caller (docx tail append). Across two shares it fails closed; S0 makes it compare shared identity. |
| Owned index, catalog, cache | `IndexedArchive` (soapberry-zip `office.rs:1703`), `PartCache` (`O/source_backed.rs:5317`) | Incidental: owned inline. Already `Sync` and read concurrently in tests. |
| `&mut self` apply and save on eager owners | docx (`D/package/package/access.rs:218-261`, `D/package/codec.rs:522`), pptx (`P/package/model.rs:550`, `P/package/codec.rs:319`, `:360`), xlsb (`B/workbook/package.rs:418-433`) | Incidental API shape. docx and xlsb already compute the new value internally and then assign it back. |
| Memo re-projection | pptx `memo_for` (`P/package/model.rs:1897-1912`); ADR 0005 2026-09-16, ADR 0032 | Calls for a hand-written `Clone` (or a fork), not for uniqueness. |
| Stated intent | change 0044: "owns the source-backed package without a `Clone` escape hatch" and "The non-clone owning editor and consuming publication call prevent retaining a second mutable publication owner" (`perf/changes/0044-pptx-source-backed-semantic-publication.md:20-21`, `:130-131`); change 0179 (`perf/changes/0179-pptx-source-catalog-reuse.md:13`); `perf/CRUD_COVERAGE.md:871` | The intent is "no second *mutable* publication owner". A shared package that never mutates, with edits carried in per-publication plans, has no mutable owner at all, so the intent still holds. |

The conclusion is the same for every OOXML leaf: ownership is demanded by
receivers (`self`, `&mut self`) and by owned fields, not by a guard that needs
a unique owner. Changing that contract is the "re-entrant publication" record
0613 asked for; §6 states the contract that replaces it.

## 4. Options

Costs below use the S1 baseline (`docs/unified-ops/results/0003/before/README.md`,
warm cache, p50, one pinned CPU on a shared host; the open rows include
detection, so they are upper bounds for a re-index from a retained reader,
*inferred*) and the leaf one-edit cycles in the performance records. The
fixture pairs are the same media corpora for docx and pptx and the same
cell-values corpus family for xlsx (*inferred* from the corpus names).

| | A: share the leaf owner | B: retain the reader, re-open when shared | C: consuming verbs + sole-owner fast path | D: A plus overlay chaining |
|---|---|---|---|---|
| Source identity, lineage | Unchanged: one opened package, one lineage | New lineage per re-open; every commit captured on the old handle becomes stale (`D/document/transaction.rs:893-896`) | Fast path unchanged; fallback inherits A or B | Unchanged |
| Exact no-ops, determinism | Unchanged; output never depends on sharing | Preserved if the re-open is source-backed and checks the captured version | Holds only if the fallback is byte-equivalent; today's xlsx fallback is not (§2.4) | Unchanged |
| Single-op budget (≤ 10 % over leaf Edit+Commit) | One `Arc` share per write (*inferred* negligible) | Adds a re-open per shared write: docx 316 µs, pptx 2,835 µs, xlsx 323 µs (file) / 62 µs (bytes). Against whole one-edit cycles of 1.71–1.86 ms (docx), 6.70–6.78 ms (pptx) and 5.31–5.42 ms (xlsx), that is ≈ 17 %, ≈ 42 % and ≈ 6 % of open+edit+publish, a larger share of Edit+Commit alone. Every Python write pays it | Rust sole owner: as A. Python: always the fallback | As A, and later writes stop paying the §5 rebase |
| 1k-op batch (≤ 2× one leaf Edit) | No conversion | One re-open per batch (pptx ≈ 2.8 µs/op at 1k) | One fallback per batch | No conversion, no rebase |
| Handle cost ∝ changed closure | Yes on eager engines and OLE snapshots, and for a first source-backed write; a chained source-backed write rebases (§5) | No: each shared write rebuilds index, catalog and cache | As A or B | Yes |
| Memory | Base shared by every handle; each handle adds its changed closure | One index, catalog and cache per re-opened handle; caches not shared, so payloads decode again | As A or B | As A |
| Rust ergonomics | `&self` verbs; `let f2 = f.set(..)?` keeps `f` | Same signatures, hidden latency spikes | `f.set(..)` moves `f`; keeping it needs `f.clone()`; a failed call loses the handle unless the error returns it (`EditorRefused`, `S/format/pptx.rs:166-172`) | As A |
| Python ergonomics | Natural; immutable handles | Every call pays the re-open | Never takes the fast path (refcount ≥ 2) | As A |
| Size (inferred) | §6: OPC 40–120 lines; xlsx 0 leaf lines; docx 5–10 for M2 (50–80 full sweep, + 60–100 eager fork); pptx 30–60 (+ 40–80 eager); xlsb 20–40 per family; OLE 0–15; ops net negative per format | Ops ≈ 100 lines per format plus a retained reader per handle; a re-open constructor per leaf with a version check | A or B plus a second verb set | A plus per-leaf overlay readers (large; follow-up) |

Cycle sources: docx `docx_source_backed_one_edit_save`
(`perf/0750-xml-audit-well-formedness-gaps.md:443`, `:739`;
`perf/0754-docx-semantic-edit-and-text-path.md:382-383`), pptx
`pptx_source_backed_one_edit_save` (`perf/0747-xlsx-publication-audit-reuse.md:314`;
`perf/0750-xml-audit-well-formedness-gaps.md:444`), xlsx
`xlsx_source_backed_cell_values_one_edit_save` medium
(`perf/0613-opc-original-audit-memo.md:196`); opens: S1 README rows
`docx_file_source_open`, `pptx_file_source_open`, `xlsx_file_open`,
`xlsx_bytes_open` (`docs/unified-ops/results/0003/before/README.md:109`, `:119`,
`:128`, `:126`).

**A — share the leaf owner.** Make one opened package shareable at O(1),
publish from a share instead of the only reference, and give eager owners a
fork that returns a new value instead of applying into `&mut self`. ADR
compliance: lineage, version and exact-source rules are untouched because
nothing is copied or mutated; the CRUD gate's "cloneable snapshots share
immutable source state" is met literally. Two mechanisms were weighed at the
OPC root; both give the same contract (§6):

- **A-clone: O(1) `Clone` for `SourceBackedPackage` (chosen).** Share the
  archive index, catalog and `PartCache` through `Arc` and derive `Clone`. There
  are two construction sites (`O/source_backed.rs:6629`, `:7053`) and one field
  move to replace (`:7884-7885`, with the copy `to_opc_package` already makes,
  `cloned_non_part_members`, `:7941`). Every share is the same opened package:
  same lineage, cache, budget and write session. Every consuming door keeps its
  signature and consumes a share; `PreparedTopology` keeps owning one. Leaves
  change least: the 16 xlsx editors are built from `package.clone()` unchanged,
  and the 58 lines in docx, pptx and xlsx that call a consuming door stay as
  they are (31 of them in xlsx); the pptx editor is
  built from a share, and the docx publishers take `&self` and clone at their
  final OPC call. The instance check at `splice.rs:1079` still passes when the
  fragment and the plan come from one value and fails closed across two shares;
  S0 should make it compare shared identity.
- **A-ref: `&self` publishers over one `Arc`-held instance (fallback).** About
  20 receivers flip (`O/source_backed.rs:9905`–`:10552` public; `:2831`,
  `:10047`, `:10561`, `:11675`–`:11751` private), plus a borrowed
  `PreparedTopology` and SVG checked topology
  (`D/source_backed/svg_lifecycle.rs:2142-2157`). Same contract, but every leaf
  editor that owns the package by value (16 in xlsx, one in pptx) must hold
  `Arc<SourceBackedPackage>` instead, which the xlsx survey sizes at
  160–320 lines. Kept as the fallback if S0 review finds the clone diff larger
  than estimated.

**B — retain the reader and re-open.** Correct only if the re-open uses the
*same* `ReadAt` (never the path, which may name a replaced file) and refuses
unless the fresh version equals the one captured at open. Neither leaf package
exposes its reader (`O/source_backed.rs:3692-3694` is private; `source_artifact()`
has no re-open), so every handle would retain a second `Arc<dyn ReadAt>`.
Its cost is a full re-index per shared write, which in Python is every write.
Rejected as a write path. It survives only inside the §5 rebase, where the
re-open is of bytes the handle just published, not of the original source.

**C — consuming verbs with a sole-owner fast path.** The fast path never fires
in Python (§1), and in Rust any live `SelectedWorksheet` or `SourceSlide`
silently disables it. Its output is deterministic only if the fallback publishes
the same bytes, which today's xlsx fallback does not (§2.4). With A as the
fallback, C adds a second verb set and saves nothing, since sharing already
costs one `Arc`. Rejected.

**D — A plus overlay chaining.** A handle after a write is "shared base +
immutable overlay" (changed parts, relationships and content types). Reads and
the next edit go through the overlay; publication writes base + overlay in one
pass. `litchi-opc` already has the substrate: `SourceTopologyPlan` and a
read-only candidate view, `EffectiveTopology::{part, parts, relationships,
content_type}` and `EffectivePart::data` (`O/source_backed.rs:2888-3037`), with
`with_prepared_topology(&self, ..)` (`:8053`); the bounded overlay doors cap a
publication at 64 parts (`:80`). What is missing is per leaf: source editors
read parts through a `(base, overlay)` view instead of the package, and patches
bind to the overlay generation as well as to lineage and version. D is the end
state that removes the rebase and meets the handle-cost budget for chained
writes. A is its prerequisite.

## 5. The batch executor

Design §3 and the goal ask for atomic batches of 1,000 or more operations,
coalesced per Edit family, within 2× of one hand-written leaf Edit doing the
same work, with per-item error attribution. Under this record:

1. **Shared to editable costs nothing, once per batch.** The executor takes the
   caller's handle by `&self` and works on a share of its state (an `Arc` clone
   of the eager snapshot, or of the source package). It never calls
   `Arc::try_unwrap`, so a batch costs the same whether or not the caller, or
   Python, still holds the handle. For comparison: B pays one re-open per batch
   (pptx 2,835 µs, about 2.8 µs per operation at 1k), today's xlsx fallback one
   complete materialization (§2.4), and C one fallback whenever any clone is
   alive.
2. **One editor per run.** Consecutive items of one (format, family, engine)
   form a run. The executor builds the run's editor once, stages every item and
   commits once. Staging is per item, so an error carries the item's index; in
   atomic mode the first error discards the run, in best-effort mode the
   executor commits the items that staged.
3. **Leaf caps split runs; they never add conversions.** The pptx source editor
   takes one operation per edit (`P/presentation/source.rs:2484-2489`) and at
   most 32 slides per batch (`:100`); xlsx source value edits take 256 cells on
   up to 64 sheets of single-worksheet workbooks (§2.4) and hyperlinks 256
   mutations (`X/hyperlinks/source.rs:17`). The docx `Edit`, the eager engines
   (xlsx `Workbook::edit`, pptx `opened::Transaction`) and the OLE transactions
   (XLS stages up to 4,096 changes, `crates/litchi-xls/src/cell_values/mod.rs:73`)
   take many operations per commit. For a large homogeneous run the executor
   picks the engine that takes the whole run in one commit, which is how the 2×
   budget is met: 1,000 cell sets are ⌈1000 / 256⌉ = 4 source commits with a
   rebase between each (inferred from the caps), or one eager commit.
4. **Chaining between runs.** Eager engines and OLE snapshots chain natively:
   the next run starts from the previous commit's snapshot. OOXML source-backed
   commits do not chain (§2.2–§2.4), so between two source-backed runs the
   executor *rebases* once: it publishes the pending commit into memory and
   re-opens those bytes, either as an `OwnedSource` package or as an eager
   owner opened from them, which keeps exact-source authorization. This is the
   only place B's re-open survives, and it re-opens bytes the batch itself
   produced. At most one rebase per family boundary, never one per item; with D
   the overlay carries across runs and the rebase disappears.
5. **Atomicity is free.** Nothing is mutated in place. A failed batch drops its
   intermediate chain and the caller's handle is untouched, which extends to
   the whole batch the property ADR 0003 states for `commit()`
   (`docs/adr/0003-snapshots-edits-and-patches.md:21-23`).
6. **No publication inside a batch.** The batch returns a new handle (base plus
   pending commit, or a rebased base); `save` publishes it.

## 6. Recommendation and staged plan

**Recommendation: option A ("share the base, never move it") now, option D
(overlay chaining) as the per-leaf follow-up.** B and C are rejected as write
paths; B's re-open survives only as the rebase of §5.

The contract that replaces change 0613's "one package therefore publishes at
most once":

- One opened package has one lineage. Any number of handles share it. A
  publication consumes a cheap share or borrows; it never needs the only
  reference.
- Each publication still checks its source (version fences, `same_source`,
  patch provenance) and authorizes itself (per-publication tokens). "Not
  published over" is a checked property (§2.1), not a structural one.
- A read-ahead window fenced by any publication stays fenced for every share.
  This affects speed only.
- Cache entries, reservations and a managed write pool live until the last
  share drops. Handles already keep the package alive for reads, so this
  retains nothing new; the pool's lifetime is reported (§9).
- Eager owners apply a commit into a new value (a fork), never into `&mut self`
  of a shared value. Exact-source authorization stays per value, as for
  `OpcPackage` today.
- The unified layer never reads a reference count. A file's bytes are a
  function of the snapshot and the operations, nothing else.

Stages follow the owner's tiers (OOXML, then OLE) and record 0002's stream
order (xlsx, docx, pptx, xlsb; then xls, ppt, doc). S0 comes first because
every OOXML stage builds on it.

| Stage | Leaf change | litchi-ops change | Tests | Size (inferred) |
|---|---|---|---|---|
| **S0** `litchi-opc` | O(1) `Clone` for `SourceBackedPackage`: `Arc`-share the archive index, catalog and `PartCache`; replace the `mem::take` at `O/source_backed.rs:7884-7885` with `cloned_non_part_members`; make the splice instance check (`splice.rs:1079`) compare shared identity; rewrite the struct doc (`:6443-6452`) and the write-session comment (`:5324-5327`) to this contract | none | a clone shares index and cache (pointer equality); publish from a clone, then read every part from the original: byte-identical; publish twice: identical bytes; exact no-op from a clone: source bytes; two clones publishing on two threads: both equal the serial output; source overwritten in place: both refuse with `SourceChanged`; managed package: one charge per decoded part across clones, everything released when the last clone drops | 40–120 lines + 100–200 test lines |
| **S1** xlsx | none required; editors take `package.clone()` | `Xlsx::editor(&self)` replaces `into_editor(self)`; delete the `try_unwrap` branch and `materialize`; promote to the eager engine through the exact source (`source_artifact` → `Workbook::from_bytes`, §2.4), once per handle; keep the tab-state editor with its pending commit | §8 on xlsx; replace `shared_handles_materialize_an_eager_editor_and_stay_readable` (`S/format/xlsx/tests.rs:539`) | ops 60–120 |
| **S2** docx | publishers take `&self` and clone the package at their final OPC call: `D/source_backed.rs:1500`, `:1570`, `:1615` for M2, the other 29 in a sweep; eager `Package` gains a fork-apply (`edit_opc_candidate` returning a new `Package`) and a clean-state `&self` save | `Docx` edits and publishes from `&self` | §8 on docx, on Litchi-written fixtures until record 0008 lifts the MCE refusal | leaf 5–10 (M2) or 50–80 (sweep), eager 60–100; ops ≈ 40 |
| **S3** pptx | `SourceBackedPresentation::editor(&self)` = `editor_seed` plus a package share (`try_into_editor` becomes a wrapper or goes); eager `Package` gains a `Clone` that re-projects its memo (`memo_for`) and `&self` fork-apply and save | `Pptx::editor(&self)`; delete `EditorRefused` and the shared refusal | §8 on pptx; replace `shared_source_handle_is_returned_readable` (`S/format/pptx/tests.rs:436`) | leaf 30–60, eager 40–80; ops −30 |
| **S4** xlsb (M2b) | a non-mutating cell-values apply that returns the new `Workbook` (wrap `apply_retaining_parse`), and the same for each in-place family as it is wired | promote through the exact source (copy the artifact, then open the eager package from those bytes) once per handle, instead of `to_workbook()` per `editor()` call | §8 on xlsb | 20–40 per family; ops ≈ 40 |
| **S5** xls, ppt, doc (M2b) | none for sharing; optional `derive(Clone)` on the commit types that hold a `ValidatedOverlayPlan` (all its fields are `Clone`) | after the first write the handle holds the leaf edit snapshot, which chains, instead of the read owner | §8 on each | 0–15 lines |
| **D** per leaf, after M2 | source editors read parts through a `(base, overlay)` view (`EffectiveTopology`) and bind patches to the overlay generation as well as to lineage and version; publish base + overlay in one pass | drop the rebase | a chained write costs in proportion to its changed closure and publishes the same bytes as the rebase path | per leaf; large |

## 7. M2 verb signatures

This record fixes the receivers and return shapes; names follow the design and
record 0007 and stay M2's to finalize.

```rust
impl File {
    // Single-op sugar: one leaf Edit and commit. The receiver is never consumed.
    pub fn set<'p>(&self, path: impl Into<PathArg<'p>>, props: impl IntoProps) -> Result<File>;
    pub fn add<'p>(&self, parent: impl Into<PathArg<'p>>, node: impl IntoNode) -> Result<Added>; // new File + new node's path
    pub fn remove<'p>(&self, path: impl Into<PathArg<'p>>) -> Result<File>;
    pub fn move_to<'p, 'q>(&self, path: impl Into<PathArg<'p>>, to: impl Into<Placement<'q>>) -> Result<File>;
    pub fn swap<'p, 'q>(&self, a: impl Into<PathArg<'p>>, b: impl Into<PathArg<'q>>) -> Result<File>;

    // Multi-op, atomic by default; the closure gets a transaction-scoped proxy.
    pub fn edit<T>(&self, f: impl FnOnce(&mut Transaction<'_>) -> Result<T>) -> Result<Commit<T>>;
    pub fn batch(&self, items: impl IntoIterator<Item = BatchItem>, mode: BatchMode) -> Result<BatchCommit>;

    // Publication of this snapshot; repeatable on any handle.
    pub fn save(&self) -> Result<()>;
    pub fn save_as(&self, path: impl AsRef<std::path::Path>) -> Result<()>;
    pub fn write_to<W: std::io::Write>(&self, sink: W) -> Result<()>;
}

/// The normalized commit shape of design §3: the new handle, its reversible
/// patch and its diagnostics, plus the closure's value.
pub struct Commit<T = ()> { /* file: File, patch: Patch, diagnostics: Diagnostics, value: T */ }
```

Rules:

- Every write verb takes `&self` and returns a new `File`, directly or inside
  `Commit`/`BatchCommit`. There is no `fn set(self, ..)`, no `&mut self` method
  on `File`, no `Arc::try_unwrap` and no reference-count branch. `File` stays
  `Clone + Send + Sync` with one pointer per variant (record 0003 D2).
- `save`, `save_as` and `write_to` take `&self` and may be called any number of
  times on any handle. On Unix, saving one handle over its source path leaves
  every other handle reading the bytes it opened (pinned descriptor; §8 T6).
- A failed verb consumes nothing; the caller's handle is as it was.
- `Transaction` is `&mut` only inside `edit`, matching ADR 0003's
  transaction-scoped proxies (`docs/adr/0003-snapshots-edits-and-patches.md:13-16`).

Python (`.pyi`), with the GIL released around each call (the handle is
`Send + Sync`):

```python
class File:
    def set(self, path: str, props: Mapping[str, object]) -> File: ...
    def add(self, parent: str, type: str, props: Mapping[str, object] | None = None, *,
            index: int | None = None, after: str | None = None, before: str | None = None) -> File: ...
    def remove(self, path: str) -> File: ...
    def move(self, path: str, *, to: str | None = None, index: int | None = None,
             after: str | None = None, before: str | None = None) -> File: ...
    def swap(self, path: str, path2: str) -> File: ...
    def apply(self, command: Mapping[str, object]) -> File: ...
    def batch(self, items: Sequence[Mapping[str, object]], *,
              mode: Literal["atomic", "best_effort"] = "atomic") -> BatchResult: ...  # .file, .results
    def save(self, path: str | os.PathLike[str] | None = None) -> None: ...
```

Every method returns a new object and leaves `self` unchanged; none takes the
inner handle out of `self`.

## 8. Tests that prove the old handle stays readable and byte-identical

Unified-layer tests, per format as each stage lands (S1–S3 OOXML, S4 xlsb, S5
OLE), each run on one real-producer fixture and one Litchi-written fixture
(the probe in §2.4 shows that Litchi-written fixtures alone can hide a
fidelity loss):

| Test | Asserts |
|---|---|
| T1 `write_keeps_the_old_handle` | `h1 = h0.set(p, v)`; `h0.get(p)` is unchanged; `h0.write_to(..)` equals the source bytes; `h1.get(p) == v` |
| T2 `write_from_a_clone_keeps_both` | `c = h0.clone(); h1 = c.set(..)`; `h0` and `c` both still write the source bytes; the base is shared, not copied (pointer equality through a test accessor) |
| T3 `output_does_not_depend_on_sharing` | `h0.set(x)` writes the same bytes with no other handle alive, with a clone alive, and with a selected worksheet or slide alive. Today's xlsx fallback fails this on every real-producer fixture of §2.4 |
| T4 `branches_are_independent` | `h1 = h0.set(a); h2 = h0.set(b)`; each equals the same single edit on a fresh open; `h0` is unchanged |
| T5 `exact_noop_through_a_shared_handle` | setting a value to itself from a shared handle writes the source bytes, including a signed source where the leaf allows a signed no-op |
| T6 `atomic_save_over_the_source_keeps_old_handles` (Unix) | open by path; `h1 = h0.set(..)`; `h1.save()`; `h0` still reads and writes the original bytes; a fresh open reads the new bytes |
| T7 `in_place_overwrite_is_typed` | overwrite the source file in place; every handle's read, write and save returns `io_error` with the leaf `SourceChanged`, never mixed output |
| T8 `concurrent_writes_from_one_handle` | N threads each run `h0.set(k_i).write_to(..)`; each output equals the serial one |
| T9 `batch_from_a_shared_handle_shares_the_base` | a 1k-op batch from a shared handle opens and materializes nothing new (OPC operation accounting or cache diagnostics) and equals the same batch from a sole handle |
| T10 `failed_write_leaves_the_handle` | an invalid item in an atomic batch returns its typed error with the item index; `h0` still writes the source bytes |

S0 adds the leaf tests listed in §6. The pytest mirror (M3) repeats T1, T2, T4
and T6: after `d1 = d0.set(..)`, `d0.get(..)` is unchanged and `d0.save(tmp)`
writes the source bytes, also with `copy.copy(d0)` alive.

## 9. Findings, risks and what is not claimed

Findings against earlier records:

- **F1** (`m0a-log.md:70`, clarification): "None of these packages can be
  cloned" is accurate for the packages, but the read owners around them,
  `SourceBackedPresentation` and both `SourceBackedWorkbook`s, are already
  cheap `Arc` clones (§2.3–§2.5). Only the OPC package inside blocks sharing.
- **F2** (record 0003 §2.2): M2 does not need `Arc::try_unwrap` or a re-open; a
  shareable package removes both.
- **F3** (`S/format/xlsx.rs:418-443`): the shared fallback breaks exact no-ops
  and byte preservation on real-producer files (§2.4), fails for managed
  handles and signed workbooks, and makes output depend on sharing. It is dead
  code today (`expect(dead_code)`) and must not ship with M2.
- **F4** (record 0002, DOCX §2): the transfer refusal has drifted to
  `D/source_backed.rs:1512-1524`.
- **F5** (record 0003 §2.3, L-DOC-1): detection yields an in-memory
  `OleFile<Cursor<Vec<u8>>>`, not a `SharedOleFile`; S5d resolved this by
  sealing that parse (`SharedOleFile::from_owned_ole_file`).

Risks:

- **Contract change.** S0 ends "one package publishes at most once". Change
  0613's original-audit memo (`perf/0613-opc-original-audit-memo.md:264-271`)
  becomes reachable; this record claims nothing about its value.
- **Retention (ADR 0005 2026-09-16, ADR 0031).** A shared package keeps its
  evictable cache, its reservations and, for a managed package that earned a
  parallel wave, its write pool and `Workers` permits
  (`O/source_backed.rs:5519-5562`; ADR 0031: "A session that holds a pool holds
  its `Workers` permits for the pool's lifetime",
  `docs/adr/0031-execution-context-budgets.md:126-127`) until the last handle
  drops.
  Python handles kept in a list keep all of it. M2 should report retained
  state per handle, as ADR 0005 asks for retention that is not a cache.
- **Concurrent managed publication.** Two shares publishing at once use one
  write session; its permits are `Mutex`-guarded, but no test runs two
  publications of one package concurrently (inferred safe; S0 adds the test).
- **Windows.** Atomic replacement over a source that another handle holds open
  is not tested (§2.1).
- **Rebase cost.** Until D lands for a leaf, a second source-backed write on a
  handle costs a publication into memory plus a re-open: O(document), outside
  the handle-cost budget. M2's benchmarks report it as a delta.
- **docx MCE refusal** (record 0008) blocks source-backed docx edits on
  Word-authored files whatever R4 decides.
- **Budget of a promotion or rebase.** Copying the exact source into memory
  (§2.4) or publishing a pending commit into memory (§5) creates bytes outside
  a managed handle's execution context. M2 must charge them to that context or
  refuse; this record does not decide which.

Not claimed: this record measured nothing except the probe's byte counts
(§2.4). Every timing is quoted from an existing record, on its own host and
base; the S1 open rows include detection. Line counts are estimates. The
fidelity-preserving promotion of §2.4 is read from the API, not run.
