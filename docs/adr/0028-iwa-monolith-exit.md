# ADR 0028: Ordered exit of the legacy IWA migration host

- Status: Accepted
- Date: 2026-08-08
- Amends: ADR 0002, ADR 0010, and the IWA record now stored as
  `0029-iwa-index-foundation.md`

## Context

Pages, Numbers, and Keynote now have concrete package owners, while the older
`litchi-iwa` crate still contains substantial editors, compatibility adapters,
examples, fuzz targets, and tests. Treating that host as another canonical
layer would let new work accumulate there and make deletion unverifiable.

The archive-free `litchi-iwa-structured` aggregation crate is not the
monolith. It currently preserves cross-format limits, ordering, text roles, and
snapshot behavior that direct packages do not yet reproduce exactly, so it
cannot be deleted merely because its name shares the prefix.

## Decision

`litchi-iwa` is the sole iWork migration host. The checked-in boundary policy
lists every direct internal dependency as ordered debt with a concrete reason
and exit condition. It has no canonical dependency allowlist. Adding an edge
without a debt record, removing an edge without removing its stale record, or
renumbering the ledger inconsistently fails the boundary check.

Residual work moves by ownership:

- physical ZIP preservation, native no-op, and package comparison belong in
  `litchi-iwa-archive`;
- archive framing and neutral metadata belong in `litchi-iwa-core`;
- object indexing and reference traversal belong in `litchi-iwa-index` and
  `litchi-iwa-graph`;
- shared raw text transformations belong in `litchi-iwa-text-wire`;
- BNC storage belongs in `litchi-numbers-wire` but is consumed privately by
  concrete format packages;
- Pages, Numbers, and Keynote topology, selectors, transactions, and native
  mutations belong in their respective concrete crates;
- root-format coordination belongs in `litchi`, without raw IDs, generated
  messages, or a compatibility re-export of the host.

No new public API may expose `litchi-iwa`, generated protobuf/Buffa values,
native object IDs, or a `raw`/`wire` compatibility module through a supported
format facade. Low-level focused crates may expose their own explicitly
unstable physical vocabulary without being glob-reexported.

## Deletion gate

The monolith is deleted only when all of the following are true:

1. No workspace or published manifest depends on `litchi-iwa`.
2. Every module, example, fuzz target, generated-schema/build path, and test
   fixture has a named focused owner.
3. Direct format packages pass semantic parity gates for all behavior being
   removed, including Numbers table order/orphans, Pages empty and fallback
   bodies, and Keynote rich storage and aggregate limits.
4. Mutation paths use selector-first concrete package transactions, preserve
   untouched bytes, and pass native application open/save/reopen tests.
5. The root facade has no re-export, feature alias, or type alias retaining the
   host.
6. The boundary policy has no `litchi-iwa` migration host or debt entry.

`litchi-iwa-structured` may remain as a neutral aggregation owner until a
separate, parity-proven rename or replacement decision is accepted.

## Consequences

The debt count may fall but never grow without an explicit architectural
change. A vertical feature may migrate before the whole format editor only
after its native gate passes. The rejected Keynote navigator-name prototype is
the counterexample: semantic readback alone did not detect native placeholder
fallback, so the concrete transaction was removed. Removing an example is not
sufficient if its behavior regresses; exact-parity tools move first, while
non-equivalent structured extraction stays in the host until its gates pass.

The duplicated ADR number on the IWA index record is now corrected: that
record is stored as `0029-iwa-index-foundation.md`. This record amends its
migration-host wording and makes concrete
format adapters, rather than `litchi-iwa::object_index`, the destination.

## Verification

`tools/check_crate_boundaries.py` must report `litchi-iwa` as a migration host
and print its ordered debt list. Source/API audits must prove that supported
Pages, Numbers, Keynote, and root facades expose no monolith or BNC bridge.
Each removed debt item also requires focused tests and, where output changes,
real Pages, Numbers, or Keynote application verification.

## 2026-08-08 Keynote storage projection progress

The concrete Keynote package now satisfies the rich-storage and aggregate-limit
portion of deletion gate 3. It builds a private sorted index over all native
objects, rejects duplicate identities, traverses only the strict reachable
document/show/slide graph, and projects referenced schema-proven type-2001 text
payloads through the bounded Buffa adapter in `litchi-iwa-text-wire`.
Incompatible native type-2022 siblings remain opaque. Body and ordinary
drawable storages retain semantic fragment ranges; unrelated messages that
happen to decode as storage are excluded. A checked format-owned profile limits
objects, slides, traversed references, decoded storages, retained fragment
ranges, and aggregate retained UTF-8 bytes,
while the original physical limits remain available and are preserved across
skip-state commits and patch application.

This does not delete the monolith. Generated Prost values still decode the
larger Keynote graph, most editor operations and compatibility tests remain in
`litchi-iwa`, and the durable patch and atomic filesystem-save gates remain
open. Ignored nested fields in those generated graph messages still rely on
the physical message ceiling rather than a complete semantic allocation
envelope. The migrated text path is nevertheless production-owned: focused tests
cover inclusive and exceeded budgets, duplicate types and identities, wrong
types and wire kinds, ambiguous ownership, false-positive payloads, concurrent
first access, native Prost/Buffa differential output, and exact reversible
skip-state behavior.

## 2026-08-08 Numbers order and orphan projection progress

The concrete Numbers package now owns both sides of the table-order contract
required by deletion gate 3. Its ordinary `Document` follows only the strict
rooted document/sheet/drawable graph and excludes detached models. Its explicit
`extract_structured_tables` compatibility projection reproduces the migration
host's archive-wide behavior: first-message classification, type-6001 pass
before legacy type-6000 pass, ascending object identity within a pass,
deduplication, and inclusion of valid detached models. A compact package index
holds one locator and at most one primary-type entry per object; object lookup
is binary-search based instead of a repeated archive scan. Checked semantic
limits cover objects, rooted sheets, rooted references, and tables while
preserving caller-selected physical limits.

Focused/legacy differential tests cover rooted order opposite global order, a
physically retained orphan, canonical-before-legacy pass order, one object
carrying both candidate types, object-vector reordering, secondary-message
exclusion, preferred malformed-model failure, ignorable malformed legacy
false positives, unrelated typed false positives, duplicate identity and
ownership rejection, and inclusive/exceeded object, sheet, rooted-reference,
and table budgets. The ordinary reader also requires exact document and sheet
types and retains only the fixture-backed type-6003 table-info compatibility
alias alongside native type 6000.

Computer Use created, saved, closed, and reopened
`/private/tmp/litchi-numbers-order-oracle-20260808.numbers` in the real Numbers
application. Its final SHA-256 is
`781181e89c655da5c92b677b9ba5c939c85379e7b33ccf10e3846fe8588f9c5b`.
The workbook has `SecondCreated` before `FirstCreated`, a non-table text box,
and tables whose rooted order is `B-only-table`, `A-new-table`,
`A-old-table`. Native Arrange ordering makes the global compatibility order
`B-only-table`, `A-old-table`, `A-new-table`. The focused example reproduced
both sequences and one materialized marker cell per table.

This closes the focused Numbers order/orphan implementation gap, not the
monolith deletion gate. The legacy aggregate structured API
(`litchi_iwa::Document::extract_structured_data`) still routes Numbers through
the migration-host adapter, mutation ownership remains largely in
`litchi-iwa`, and the larger Numbers table graph still uses generated Prost
messages. The host adapter can be removed only after a source-owning aggregate
coordinator consumes the focused projection and the remaining compatibility
tests move to focused owners.

## 2026-08-08 Numbers compatibility-ingress hardening

The focused format owner now exposes
`compatibility_tables_from_bytes[_with_options]` as a distinct global
projection ingress. It validates the same immutable byte snapshot, builds the
compact object index, and extracts compatibility tables without first
constructing the strict rooted workbook. This preserves the deliberate
rooted/global distinction for detached tables and malformed unrelated rooted
topology while giving a future source-owning aggregate coordinator a direct
format API.

Both direct package and compatibility ingress prove an unambiguous Numbers
root from the unique canonical type-1 payload through `litchi-iwa-detect`
before TN decoding; application-shaped siblings cannot mask it. The focused
index now rejects null object identifier zero in addition to missing and
duplicate identities. Global candidate extraction decodes only the payload
that supplied the object's primary classification, so a primary type-6000
metadata object cannot be promoted by a secondary type-6001 payload; duplicate
canonical or legacy model payloads fail closed. The caller's table ceiling is
intentionally charged before decoding another canonical candidate, preventing
malformed over-budget input from forcing an otherwise disallowed model
allocation.

Formula-reference enrichment is no longer built by every table extractor. It
is initialized only after a non-empty formula sidecar is selected, resolves
objects through the compact index instead of repeated archive scans, uses
fallible map growth and shared table/sheet names, and charges the configurable
reference budget only for unique source-derived retained entries. Discovery
work, encoded category bytes, and cumulative source text use fixed package-wide
caps; category depth has a fixed per-tree cap. Each category payload is
schema-preflighted before a private Buffa projection validates an empty node
envelope plus UUID and scalar wrappers. Recursive children and CellValue
branches stream from the
preflighted source rather than entering generated repeated-fragment storage;
an O(depth) iterator stack walks children and ignored native fields stay
opaque. Filesystem open is nonblocking on Unix, rejects non-regular descriptors,
obtains metadata from the opened descriptor, fills the bounded destination
buffer through the standard spare-capacity reader path, and verifies the
descriptor version afterward. This closes the earlier stat/open race and
rejects observable in-place mutation. Protobuf failures are mapped to a
Numbers-owned, content-free semantic location and retain no generated decoder
source.

This is a cutover prerequisite, not the cutover. Reopening a legacy
`litchi_iwa::Document` through `Package` would still break in-memory and
directory-backed inputs, violate immutable snapshot semantics, duplicate
physical parsing, and impose strict rooted failures on the historical global
API. A shared immutable catalog/source coordinator and package-wide
compatibility budgets for sidecars and their decoded allocations remain
required before deleting `structured/numbers.rs`. The remaining table model,
formula-owner, sidecar, and AST Prost decoders still need focused pre-decode
envelopes or projections, and the broad public text projection is not yet
covered by the table-projection text budget.

## 2026-08-08 Numbers aggregate projection and formula-render hardening

The focused rooted and global table projections now share caller-selected,
package-wide budgets for materialized cells and retained semantic text. A
table charges its dimensions before allocating cell offsets, and sheet names,
table names, retained cell text, rich text, formula errors, and rendered
formula text charge one aggregate output budget. Compatibility candidates use
a transactional budget snapshot: a malformed speculative legacy candidate
does not consume retained cell or text capacity, while a successfully
published table commits those charges. Formula AST work is monotonic across
both successful and rejected candidates so hostile speculation cannot receive
a fresh CPU allowance. Existing per-table cell limits remain as a second,
narrower bound.

Formula rendering no longer recursively builds and copies an intermediate
`String` at every AST node. It constructs an arena of nodes and string parts,
charges a shared AST-work budget and a thunk-depth budget, computes the final
size with checked arithmetic, reserves exactly once, and emits iteratively.
Differential tests compare this renderer with the former implementation.
A 4,096-value skewed concatenation proves linear arena growth, and exact-limit
and one-over tests cover work, depth, output text, and aggregate cells. Formula
metadata is still initialized after selection of a non-empty formula sidecar;
deferring it until an individual rendered formula requires cross-table or
category resolution remains open.

Application classification also fails closed when a package without the
canonical Numbers type-1 root instead has an unambiguous Pages- or
Keynote-shaped root. Synthetic coverage and the checked-in native Pages and
Keynote fixtures return `NotNumbers` through both direct and compatibility
ingress. A compiler-backed CI ratchet now marks `litchi-numbers-wire` as a
private dependency and denies exported-private-dependency signatures and
blanket conversions. The public Prost error wrapper and public wire/comment
conversions were removed; malformed generated payloads expose only a
format-owned semantic path.

Computer Use authored, saved, closed, and reopened
`/private/tmp/litchi-numbers-formula-richtext-native-20260808.numbers` in the
real Numbers application. The workbook contains a numeric input, a stored
`SUM` formula whose reopened result is `323`, formatted `Café` text in a
cell, and a two-line formatted text box. Numbers reopened it without a repair
or conversion prompt. Its SHA-256 is
`80deb7b87df27f58b26e6f247acee9d1fc6dcd3d268e85046c3efc16070b2edf`.
The focused example reads one rooted sheet, one global compatibility table,
and six materialized cells; reading does not change the hash.

This closes the aggregate cell/text and formula-render allocation gaps for
materialized table projection, not the complete Numbers cutover. Table-model,
table-data, formula-table, and AST payloads still enter eager Prost decoders;
sidecar work and decoded-memory envelopes remain incomplete. The public API
also still carries archive/common physical debt and low-level formula
identifiers outside the new wire-dependency ratchet. The root structured
coordinator and removal of the host Numbers adapter remain separate work.

## 2026-08-08 Keynote root projection and host setter deletion

The concrete Keynote package no longer fully Prost-decodes the root
`KN.DocumentArchive`. A narrow generated Buffa lazy view selects only the show
reference, while a Keynote-owned wire preflight requires unique canonical
fields, validates all currently known reference scalars, and deliberately
keeps the ignored document-super envelope opaque. The deferred reference is
forced before publication. The generated closure is provenance-checked,
forbids unknown retention and element-memory support, and is held to five
generated files and 64 KiB. Differential and hostile-input tests cover native
payload parity, malformed references, missing and duplicate identifiers, and
a 256 KiB opaque super payload.

The migration host's `KeynoteEditor::set_slide_skipped` was deleted together
with its duplicate legacy assertions. The focused Keynote transaction remains
the sole owner of this mutation and already proves reversible patching,
unknown-wire preservation, and native application behavior. This removes one
concrete host operation, but slide nodes, slides, builds, shapes, notes, and
most Keynote editor workflows still use larger generated graphs or remain in
`litchi-iwa`.

## 2026-08-08 legacy Numbers BNC parity

The remaining migration-host table extractor now interprets legacy BNC
type-9 cells through the same private stored-value model as the focused
Numbers package. Rich text wins over string, which wins over numeric; formula,
cached-scalar, error, and comment precedence follows the focused union
semantics. Duplicate v5 flag walking and decimal conversion were removed.
Unit tests cover numeric and precedence cases, and a whole-package
focused-versus-legacy differential fixture locks the behavior while this host
reader still exists.

## 2026-08-08 immutable source-catalog prerequisite

`litchi-iwa-archive` now owns a `SourceCatalog` that binds one authoritative
immutable byte snapshot to both the physical/logical package catalog and its
deterministically ordered IWA components. Borrowed ingress copies the source
once, shared `Arc<[u8]>` ingress retains the exact allocation, and positional
ingress inherits the existing bounded source-version check. Direct ZIP sources
carry `ExactZip` provenance; normalized nested `Index.zip` sources carry
`LegacyZip` provenance and cannot be mislabeled as exact preserve-mode input.
Component decoding consumes the package catalog's already-decoded logical
members, so it does not reopen the ZIP or decompress the same ZIP member a
second time. Operation storage remains opaque and unsupported compression on
an IWA member fails closed.

Focused Pages and Keynote packages now retain this shared snapshot. Pages
extracts metadata and native components from the same catalog instead of
performing two complete package ingresses. Keynote classifies the already
parsed component catalog under the caller's original physical profile, reads
metadata from the retained physical catalog, and reuses that catalog for
slide-state edits; detection, metadata, and editing no longer reopen its source
through separate ZIP parsers. Catalog-based detection has differential coverage
against byte-based detection for all three application roots. Unit evidence
also counts one physical catalog construction for a direct ZIP and exactly two
for a genuinely nested legacy ZIP, while proving shared-source allocation
identity and component parity with component-only ingress.

Computer Use reopened the checked-in native Pages, Numbers, and Keynote
fixtures in their respective applications without repair or conversion UI and
closed them without saving. Their hashes and visible semantic markers remained
the documented native oracle. Focused native fixture tests continue to cover
the migrated Pages and Keynote readers.

This is the source-owning aggregate prerequisite, not the root coordinator or
the monolith cutover. The catalog still eagerly materializes decoded package
members and neutral IWA archives; it is not a claim that the full aggregate
graph is Buffa-lazy. Numbers compatibility projection has not yet accepted a
shared catalog, directory bundles and mutable `EntryStore` snapshots still lack
a frozen logical-entry adapter, and Pages/Keynote aggregate contract differences
remain unresolved. The next cutover stage must add those handoffs and
role-aware root parity tests before deleting any host structured adapter.

## 2026-08-08 limits-preserving prepared-source handoff

The focused detector now owns an opaque, single-use `PreparedSource`. It binds
application classification to the same immutable `SourceCatalog` that one
selected format owner consumes, so root coordination no longer needs a direct
archive dependency or a detect-then-reparse byte path. Borrowed ingress copies
once into immutable storage and shared `Arc<[u8]>` ingress retains allocation
identity. Non-ZIP and unrecognized inputs remain unclaimed rather than being
coerced into an iWork format.

`SourceCatalog` now records and exposes the validated physical `Limits` that
authorized ZIP, Snappy, and neutral IWA parsing. Pages and Keynote derive their
physical and text assumptions from that retained profile; a handoff cannot
silently supply a weaker second profile after validation. Their explicitly
unstable constructors, and Numbers' corresponding global compatibility
projection, are enabled only by `internal-iwork-source`. The root `iwork`
feature forwards those private integration features, but no supported format
facade returns a prepared source, catalog, archive, protobuf value, or raw
identifier.

Numbers consumes the prepared catalog directly into its existing compact index
and global compatibility projector. It deliberately does not construct the
strict rooted workbook, preserving detached/orphan table behavior and the
established global source order. Pages and Keynote consume the catalog by move;
Keynote retains the authoritative source allocation for exact no-op and
preserve-mode editing.

The archive-free `litchi-iwa-structured` owner can now retain a Pages semantic
`Document` or Keynote semantic `Document` directly while preserving its public
slice and text-role APIs. Aggregate construction validates the same count,
canonical-position, and text budgets but does not clone a section, slide,
storage, run, build, transition, or string. Numbers remains an owned `Vec<Table>`
because its required global compatibility projection is the first unavoidable
materialization and already transfers that allocation without another copy.

Focused tests prove retained physical profiles, shared-source pointer identity,
direct-versus-handoff semantic parity for all three native fixtures, preserved
Numbers global semantics, and Pages/Keynote document pointer identity across the
structured boundary. This completes the no-reparse and no-deep-clone handoff
foundation. It is still not the supported root coordinator or permission to
delete a migration-host adapter: directory/`EntryStore` frozen sources,
root-owned errors and value wrappers, aggregate cache policy, and role-aware
root parity remain required.

Computer Use reopened the native Pages, Numbers, and Keynote fixtures for this
handoff gate. Pages exposed one body containing the three expected lines;
Numbers exposed `Table 1` as 22 rows by 7 columns with the expected B2 text and
B3 numeric value; Keynote exposed separate title, body, and date text boxes.
No application presented repair or conversion UI. Each document was closed
without saving and all three SHA-256 hashes remained unchanged.

## 2026-08-08 root-owned immutable structured coordinator

The root `litchi` package now owns the supported read-only cross-format API at
`litchi::iwork`. Borrowed bytes and caller-owned immutable shared bytes enter
one finite physical profile, form one opaque `PreparedSource`, and are
classified exactly once. That single-use value is consumed by precisely one
selected Pages, Keynote, or Numbers owner. A successful root `Document` is
eagerly decoded and aggregate-validated, so its `snapshot`, table, slide, and
section operations are infallible views rather than deferred parse points.

The public boundary is facade-owned. `Format`, `Options`, physical and
semantic limits, content-free errors, `Document`, `Snapshot`, lifetime-free
table/slide/section handles, borrowed text roles, and Numbers cell values are
root types. A rustdoc-JSON gate rejects public lower iWork crates, concrete
format types, Buffa/Prost types, archive/catalog/prepared capabilities, and raw
identifier vocabulary. The root package has a canonical edge to the neutral
archive-free structured owner, but still has no edge to `litchi-iwa` or the
archive crate.

The selected semantic contracts are deliberately format-owned:

- Pages uses the focused semantic document. An empty root therefore has zero
  sections; bounded fallback bodies, native section names, and UTF-16 section
  boundaries remain authoritative instead of preserving the narrower host
  projection.
- Keynote preserves navigator name separately from visible title, retains
  skip/build/transition state on its lifetime-free slide handle, and orders
  root text as title, ordinary content, additional rich text, then notes.
- Numbers consumes the global compatibility projection, including detached or
  orphan tables and its established candidate ordering. It never substitutes
  the stricter rooted workbook constructor.

Pages and Keynote transfer their cheaply shared semantic documents into the
neutral aggregate without cloning their contained values. Numbers transfers
the first unavoidable global `Vec<Table>` without a second materialization.
After aggregate construction, the concrete package and physical source are
dropped. Native fixture tests use `Weak<[u8]>` to prove that the original
package allocation is released while cloned root handles remain usable.
`SourceCatalog::into_components` also releases the physical Numbers catalog
before compatibility-table projection. Aggregate decompressed IWA retention is
now charged across all component streams against the existing total expanded
byte profile, with exact and one-over coverage for both component ingress
routes.

Root tests cover all three native fixtures, unrecognized input, format
isolation, typed Numbers cells, role-aware text order, `Send + Sync`, cheap
snapshot/handle lifetime, and exact versus one-over input and semantic text
ceilings. The three native files were reopened through Computer Use in Pages,
Numbers, and Keynote without repair or conversion UI, closed without saving,
and retained their documented hashes. This evidence proves provenance and
nonmutation of those fixtures only.

This amendment does not authorize deletion of the migration-host structured
adapter. Frozen directory bundles and mutable logical-entry snapshots still
lack a source-owning root route; host parity/property/concurrency/fuzz/example
ownership is incomplete; Numbers retains eager Prost/sidecar allocation debt;
and concrete editing, atomic saving, and native save/reopen gates remain.
Consequently `litchi-iwa::Document::extract_structured_data`, its structured
modules, dependency, and boundary-debt entry remain until those gates close.
No performance gain, complete Buffa laziness, directory parity, or resave
fidelity is inferred from the new dependency shape.

## 2026-08-08 frozen path ingress and semantic-only projection

The supported root coordinator now accepts filesystem paths through
`litchi::iwork::Document::open[_with_options]`. A regular file is opened once
with no-follow and nonblocking flags on Unix, bounded before publication, and
checked for descriptor/path identity, type, length, modification, and change
metadata around the read. A directory is captured once by the archive-owned
`FrozenDirectoryBundle`. Exactly one direct `Index.zip` or loose `Index/`
representation is allowed; symlinks, special nodes, nested loose directories,
dual representations, marker conflicts, unstable manifests, and observable
source replacement fail closed. Component order is normalized and every
physical, aggregate-IWA, and allocation limit is applied before publishing the
snapshot. The detector classifies the same retained components subsequently
consumed by the chosen format owner and preserves content-free limit,
allocation, encryption, invalid-profile, and source-change categories through
the root error boundary.

Directory provenance is intentionally narrower than ZIP provenance. The
frozen value owns only the semantic index representation and application-marker
evidence. `Metadata/`, `Data/`, previews, and unknown root sidecars are outside
that adapter. It cannot expose exact package bytes, enter preserve-mode edits,
or claim directory reassembly fidelity. The filesystem cannot provide a
cross-file atomic snapshot without an external filesystem snapshot or lock;
the adapter instead rejects every observable change during its bounded capture
and performs no later filesystem reads.

All three root branches now consume a component-only semantic handoff. Pages
releases the package catalog before root/reference validation and applies the
root section and text ceilings during its first semantic construction. Numbers
shares the retained component catalog through `Arc` and preserves the global
orphan-compatible table projection. Keynote uses a private component-backed
semantic decoder, so directory reads do not construct package metadata or edit
state. Its lazy semantic cache now uses fallible single-flight initialization:
concurrent first readers perform one decode, failures remain retryable, and a
prepared source bypasses redundant format classification. The package-oriented
Pages and Keynote constructors remain available for exact ZIP metadata and
editing, but safely reject directory-backed prepared sources.

Native directory oracles were produced from disposable copies of the three
checked-in fixtures in the real Pages, Numbers, and Keynote applications using
the application's Package file type. Each package directory was saved, closed,
reopened from Recents, checked for the same visible Pages lines, Numbers table
and B2/B3 values, or Keynote title/body/date, and closed without another save.
No repair or conversion UI appeared. The 46 regular members and their hashes
are checked in under `test-data/iwork/directory`; the original ZIP fixture
hashes remained unchanged. Root integration tests prove ZIP/directory semantic
parity for all three formats, exact and one-under directory input ceilings,
typed missing/link/special/mixed-source failures, and stable semantic handles
after the captured directory has been removed.

Root-owned migration infrastructure also advanced. `crates/litchi` now owns a
feature-gated `read_iwork` example that uses the bounded path API and a fuzz
package whose only parser dependency is `litchi` with `iwork`; the legacy fuzz
target no longer calls the host aggregate method. The public API gate compiles
thread/lifetime assertions and feature isolation in addition to checking
rustdoc JSON. It intentionally does not use `--locked`, because this library
workspace excludes `Cargo.lock` from version control and the gate must work in
a clean checkout. The root fuzz harness compiles and records native seed hashes,
but a sanitizer campaign is not claimed when `cargo-fuzz` is unavailable.

`litchi-iwa-package::EntryStore` now has cheap immutable `freeze` and `snapshot`
views with copy-on-write isolation, deterministic positions, and `Send + Sync`
coverage. That is only the storage seam: it has not yet been admitted through
one root prepared-source coordinator or proven against all host logical-entry
behavior. The monolith structured adapter and its dependency debt therefore
remain. Deletion still requires the frozen logical-entry route, retained
host-versus-focused parity oracles for every removed behavior, completed fuzz
execution, and migration of the remaining editors/tests/examples that depend
on broader host semantics. This slice does not claim complete Buffa conversion;
the remaining focused Prost graph decoders stay tracked migration work.

## 2026-08-08 validated logical ingress and capability-anchored directories

The archive boundary now owns `LogicalSourceCatalog`, the admission point for
an immutable `litchi-iwa-package::FrozenEntryStore`. Construction performs a
complete validation pass over every entry before decoding any IWA stream:
entry count, exact portable name, individual and aggregate name metadata,
individual and aggregate payload bytes, encryption markers, and any basename
equal to `Index.zip` all fail through typed physical categories. Because the
input is already a logical package, the physical `max_input_bytes` ceiling is
not reinterpreted as a payload ceiling; entry and expanded-byte limits remain
authoritative. The same frozen store is retained through one component
classification and is dropped before format-owned semantic decoding. This
route never synthesizes ZIP bytes, claims exact-save provenance, or produces a
`SourceCatalog`.

The detector exposes this only through doc-hidden prepared-source integration
methods. The supported root API deliberately gains no entry-store constructor:
the migration host did not expose a direct logical-entry API, and publishing
one would leak package names and physical capabilities into the semantic
facade. The root rustdoc gate now rejects `litchi_iwa_package` in addition to
the other implementation crates. Tests cover copy-on-write isolation without
payload copying, direct-ZIP classification parity for all three applications,
operation-log exclusion, exact and one-over logical limits, unsafe names,
encryption, and unexpanded nested indexes.

Directory capture is now capability-anchored on Unix. The final bundle root is
opened with no-follow semantics and pinned; `Index`, its manifest, and every
member are acquired relative to retained descriptors. Descriptor identity,
node type, byte length, manifest contents, application-marker evidence, and
the selected `Index.zip` or loose `Index/` representation are revalidated after
component parsing and before publication. Root and `Metadata` encryption
markers, loose-index encryption markers, nonportable basenames, exact read
length, and the aggregate loose payload against both input and expanded-byte
ceilings are enforced. Replacing an ancestor pathname after the root is open,
the root pathname itself, `Index`, or an individual member cannot redirect a
published snapshot. A pre-existing ancestor symlink is still resolved by the
initial operating-system path lookup and is documented as such.

Non-Unix capture retains the path-based identity/revalidation fallback. It now
shares the encryption, aggregate accounting, portable-name, read-length, and
post-parse checks, but it does not claim the same adversarial replacement
resistance as descriptor-relative Unix acquisition. Cross-file atomicity also
remains unavailable without a filesystem snapshot or external lock. These are
explicit portability limits, not inferred security guarantees.

The archive-free aggregate corrected two semantic invariants. Keynote text is
now consistently ordered as title, ordinary content, additional rich storage,
then notes, matching the root and focused leaf contracts. Retained text budgets
now include Keynote navigator names and Pages section names even though those
identity strings are not emitted by `iter_text`. Exact and one-under tests
cover the additive budget, storage/notes order, empty storage filtering, and
leaf/aggregate parity. The root error vocabulary also distinguishes objects,
sheets, references, text storages/fragments, payload bytes, fields, and nesting
depth; known Pages, Keynote, and Numbers limits map exactly, while invalid
aggregate positions report validation invariants. Nested Numbers common-error
classification remains a leaf-owned follow-up rather than introducing a root
dependency on `litchi-iwa-common`.

Migration ownership moved forward without deleting behavior. The obsolete
host `read_iwork` example is superseded by the bounded root example, the
Numbers structured-extraction example now lives in `litchi-numbers`, and the
host-only `once_cell` use is a development dependency. The migration host's
detector compatibility conversion now handles every expanded detector category
and retains a future-proof fallback; all 1,479 host library tests pass.

Computer Use reopened the checked-in directory fixtures in Pages, Numbers,
and Keynote. The expected Pages three-line body, Numbers 22-by-7 table with its
text and numeric marker cells, and Keynote title/body/date were visible without
repair or conversion UI. The applications nevertheless rewrote each
`Index.zip`, `Metadata/DocumentIdentifier`, and `Metadata/Properties.plist`
merely by opening the packages. The manifest gate detected all nine changes;
the exact tracked bytes were restored and every checked-in member hash passed.
This is visual compatibility evidence, not a native nonmutation claim, and
future direct application checks must use disposable copies.

The root fuzz target compiles offline, but `cargo-fuzz` is not installed and no
sanitizer campaign was executed. At the time of this 2026-08-08 gate, the host
structured adapter, its dependency, and all 17 recorded monolith debt edges
remained; later debt-retirement amendments supersede that historical count.
This amendment does not authorize monolith deletion, claim complete Buffa
laziness, or infer edit/resave fidelity from the new source routes.

## 2026-08-08 Keynote Show/SlideTree and slide-order ownership

The concrete Keynote owner now removes two more eager generated-graph uses from
its show boundary. A derived private Buffa lazy projection covers the supported
`KN.ShowArchive` settings and required envelopes after a schema-directed wire
preflight. The embedded `KN.SlideTreeArchive` is routed manually: its ordered
slide references are streamed from validated source fields so Buffa never
builds an attacker-width nested repeated-message index. Required reference
identifiers, known optional reference scalars, canonical wire framing,
required-envelope uniqueness, setting presence, finite semantic values, slide
and reference budgets, and every deferred value used for publication are
checked before the semantic `Show` is visible. Generated Buffa/Prost values and
the native slide tree remain private; accepted raw source bytes retain
preservation authority.

`Package::edit_slide_order()` now stages one selector-first move in a separate
`SlideOrderEdit`. Exact navigator names and checked semantic source positions
are accepted; the typed destination is the final zero-based position in the
base list and must be less than its slide count. A same-position move shares
the source allocation and exact bytes. A real move reorders complete raw
slide-reference field records, including each encoded key, encoded length, and
nested reference payload, preserving unknown and deprecated fields with their
slides. It then reopens the complete package under its retained `ReadOptions`
and verifies semantic order. `SlideOrderCommit`, `SlideOrderDiagnostics`,
`SlideOrderPatch`, `SlideOrderError`, `SlideOrderLimitKind`,
`Package::apply_slide_order`, and the inverse patch keep native identifiers and
component names private and require exact source bytes for publication.

The migration host's `KeynoteEditor::move_slide`, its raw-index-only example,
and its move-specific compatibility assertions are retired after their focused
equivalents take ownership. This is a vertical behavior move, not permission to
drop slide creation, duplication, deletion, show settings, or the larger
Keynote editor graph. Those paths have distinct component-registration,
allocation, dependency-disposition, and reclamation contracts.

Acceptance evidence was executed rather than inferred:

- **Rust:** the protobuf crate passed 38 unit tests; Keynote passed 67 unit, 37
  integration, and 2 doctests; the migration host passed all 1,478 library
  tests; the direct Keynote root facade passed 2 tests; the aggregate iWork
  facade passed 8; and the structured owner passed 12. Warning-denied Clippy
  passed for every protobuf target, every Keynote production/library/example
  target, and the full slide-order test target. Formatting and diff checks
  passed. The host-versus-focused native differential produced `B/C/A` in both
  readers and byte-identical extracted `Index/Document.iwa` output. The
  focused writer additionally retained untouched ZIP metadata that the legacy
  host normalized. The unrelated host-wide examples check remains blocked by
  a pre-existing Numbers example that accesses a private raw sheet ID; host
  library compilation and tests pass.
- **Generated boundary:** the derived schema is 1,682 bytes; Buffa 0.9.1 emits
  exactly five files/138,661 bytes and no generated repeated view. The build
  checks canonical schema declarations and handwritten route constants.
  Public-API audit passes, and the boundary checker reports 63 packages, 224
  internal declarations, and the expected 17 ordered migration debts.
- **Native Keynote:** Keynote 14.4 (7043.0.93) authored disposable `A/B/C`
  source
  `/private/tmp/litchi-keynote-order-oracle-20260808.B6vCko/source-abc.key`
  (`49c7ee349cddb9fcd4671b7cd36c90008a76e457311cd3bb70d4b765f217b3df`).
  The focused move `0 -> 2` produced `litchi-moved-bca.key`
  (`62960a755535fd719bffa53f6f9e9f6126fa22d2ae50c3b543e24f926da07779`).
  Keynote opened it without repair, recovery, or conversion and displayed
  `B/C/A`; native Save As produced `keynote-resaved-bca.key`
  (`81f2e6010f68504fc58b2c948604f05f3651e3252ddba10c98b7eee29aed16e9`),
  whose close/reopen navigator and focused reverse read both remained `B/C/A`.
  The public inverse restored the exact source hash. All ZIP payloads except
  `Index/Document.iwa` are identical; its focused/legacy output hash is
  `9ecd2426425491053898658f5b7584d0633b30d3a3b020bf226d397f7693d310`.
  Decompressed comparison reports only Show object 2652385 changed and its
  archive metadata unchanged. ADR 0008 records the exact commands and expanded
  evidence.

No latency, RSS, allocation-performance, fuzz, or sanitizer result is claimed.
At the time of this 2026-08-08 gate, all 17 ordered host dependency debts
remained; later debt-retirement amendments supersede that historical count.
Slide nodes, slides, builds, shapes, notes, tables, charts, media, other
mutation paths, and portions of semantic graph projection still use the
migration host and/or generated Prost values. Protobuf groups remain
transactionally fail-closed at shared package preflight. The unavailable
sanitizer campaign, missing aggregate transaction peak-memory option, durable
JSON patch envelope, atomic filesystem save, and remaining examples/tests/fuzz
targets keep the monolith deletion gate open.

## 2026-08-08 focused Keynote settings and direct graph-edge retirement

The preceding Show/SlideTree section's 17-debt count and its statement that all
show-settings mutation remains in the host are superseded by this amendment.
`litchi-keynote::Package::show_settings()` now reads validated presentation
settings directly from the retained Show payload. It validates the complete
known Show and SlideTree envelope and the slide-reference ceiling, then forces
only the private Buffa size/scalar projection. It does not initialize the full
semantic slide cache or retain slide-node identifiers. Buffa does not retain
unknown content; accepted raw source records remain authoritative.

For a present Show in an exact package, `edit_show_settings()` stages the
archive-free `Settings` value and publishes a changed candidate only after one
owning IWA component is rewritten, the complete package is reopened under the
retained `ReadOptions`, and the focused reader reproduces every requested
setting. Exact no-ops share the original source allocation and bytes. A null
root show reads as `Settings::default()` and supports only that exact no-op,
because this transaction does not allocate an object or register a component.
The reversible patch retains exact source/target artifacts privately and uses
exact bytes, rather than its public diagnostic fingerprint, for conflict
authorization.

The preservation boundary is explicit. Untouched ZIP entries and raw ZIP
records, non-setting Show field records, nested unknown Size records, and the
source snapshot remain exact, including unchanged encoded field keys and
length headers. The changed Show message's effective type and length, its
`MessageInfo`, and the enclosing framing required by a changed length are part
of the intended mutation closure. They are not claimed as unchanged metadata.
Changed legacy nested-`Index.zip` sources return typed
`UnsupportedSource`; silently flattening them would violate preserve-by-default.
The migration host therefore retains its normalizing compatibility method,
example, and assertions. This is focused exact-source ownership, not full host
show-settings retirement.

Ordered dependency debt 007 is independently retired. The host reference
adapter inserts authoritative and fallback edges directly into
`litchi-iwa-index::IndexBuilder`, and other host users obtain `ObjectId` and
immutable graph snapshots through the index owner's reexports. Strict builder
insertion still rejects duplicate references, while the new adapter-specific
idempotent insertion preserves native duplicate-deduplication behavior. Null
handling, authoritative-list suppression of fallback, deterministic ordering,
and missing-target visibility remain unchanged. `litchi-iwa-index` still owns
the canonical graph dependency; this does not claim migration of the remaining
graph-backed editors. The ledger now contains 16 debts, with identity 007
absent and identities 008 through 017 unchanged.

Executed Rust evidence includes 31 `litchi-iwa-core` tests, 38
`litchi-iwa-protos` tests, 9 `litchi-iwa-index` tests, all 11 focused
`show_settings` integration tests, 1,479 migration-host library tests, 3 direct
root Keynote facade tests, and 3 Keynote doctests. The final Keynote
all-features/all-targets run passed 68 library tests and 48 integration tests
across eight integration binaries. Scoped warning-denied Clippy
passed for the changed Keynote, protobuf, index, example, and focused test
targets. A full Keynote dependency Clippy traversal remains blocked by 88
pre-existing `litchi-core` ARM SIMD lint failures and is not represented as a
passing gate. Formatting, diff checks, supported rustdoc public API checks,
and the boundary checker passed; the latter reports 63 packages, 223 internal dependency
declarations, and exactly 16 ordered debts.

Computer Use verified the exact-source writer in Apple Keynote 14.4
(7043.0.93). The source
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/source.key` has
SHA-256
`f3adcde9315b6df580805bcb63c995cc1e1ef569a4befa06a102485e13c883b2`.
The pristine Rust candidate was reproduced after the final code gate as
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/final-rust-reproduced.key`
with SHA-256
`c8364bb21713892f6c3c5dfb37207f8d293f48010ad16c1ff3da0547ea9f0644`;
its public inverse reproduced the exact source hash. These are the same
candidate bytes originally presented to Keynote. The opened working path after
Keynote's in-place autosave is
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/final-self-playing.key`
with SHA-256
`a106977db366e794be087a87ddfd874e7af3c26fa84d9fb5d573ca74efec739a`.
Keynote opened and automatically played it without a repair, recovery, or
conversion prompt. The inspector reported Self-Playing, loop enabled,
automatic play on open enabled, 1920-by-1080 Widescreen, a five-second
transition delay, and a two-second build delay.

Native Save As, close, and reopen produced
`/private/tmp/litchi-keynote-show-settings-20260808.g4cipH/final-keynote-resaved.key`
with SHA-256
`a9109add346eb26c8a9cb6f7db7e6bd6f1a6366a6ba1c9d073ac1c7c64bc6857`.
The focused reader recovered the inspected settings; focused no-op and inverse
outputs over that final native artifact remained byte-identical to the
`a9109add...` artifact. Before native application normalization, applying the
Rust transaction's public inverse restored the exact original
`f3adcde9...` source. ZIP entry names were unchanged and only
`Index/Document.iwa` content changed in the pristine `c8364bb2...`
Rust-authored package.

No O(1), single-pass, latency, RSS, allocation-performance, fuzz, sanitizer,
or complete Buffa-laziness claim is made. The legacy settings normalization
path, most Keynote editors and generated Prost graph paths, durable patch
serialization, an aggregate transaction peak-memory option, atomic filesystem
save, remaining examples/tests/fuzz ownership, and all 16 remaining host debts
keep the monolith deletion gate open.

## 2026-08-08 focused Pages section-name ownership

`litchi-pages` now owns selector-first replacement and removal of existing
section names for exact package sources. The transaction resolves a semantic
position, preserves absent versus explicitly empty presence, rewrites only
native field 26 in one selected section message, preserves the full object
header with the bounded shared core helper, reassembles one component, reopens
the complete candidate under retained limits, and verifies the published
section projection. Generated Buffa and Prost values, raw IDs, member names,
and wire records remain private; validated raw records are the preservation
authority for this mutation.

Exact no-ops—including legacy nested-`Index.zip` inputs—share the original
source allocation. Changed legacy sources return typed `UnsupportedSource`.
The host's raw-ID rename example is removed in favor of the focused semantic
example, but its `PagesEditor::set_section_name` normalizing compatibility path
remains until legacy mutation has an explicit preservation-safe owner. No
manifest debt is removed, so all 16 current ordered debts remain.

Apple Pages 14.4 opened the Rust artifact without repair or conversion,
preserved the body markers, completed native Save As/close/reopen, and produced
a native-resaved artifact whose expected section name reverse-read as an exact
byte-identical no-op. The public inverse restored the pre-application source
artifact exactly. This evidence advances one focused Pages exit condition; it
does not satisfy durable patch serialization, atomic save, aggregate peak
memory, fuzz/sanitizer, remaining editor/test/example ownership, or complete
host deletion.

## 2026-08-08 focused Pages section-pagination ownership

`litchi-pages` now owns exact-source read, edit, reversible patch application,
and inverse replay for `TP.SectionArchive` pagination fields 20--22. The public
surface selects an existing section by exact semantic name or checked position
and exchanges only the presence-preserving `Pagination` value. Native object
identifiers, component names, protobuf messages, wire records, and exact patch
artifacts stay private. The private Buffa sidecar is a bounded lazy scalar
projection; validated caller-owned records remain the preservation and rewrite
authority.

Changed edits preserve unknown section fields and the complete IWA object
header, mutate one package member, fully reopen the candidate, and verify the
semantic result. Exact no-ops share the source allocation even for legacy
nested packages, while changed legacy sources are refused. The host raw-ID
pagination example is removed in favor of the focused selector-first example.
The host settings/background compatibility writers remain, but now use the
bounded header-preserving message replacement helper instead of replacing the
message and silently rebuilding its header metadata.

Apple Pages 14.4 opened the Rust-authored right-page/restart-at-7 artifact
without repair or conversion, retained the fixture content, displayed page 7
and `Start at: 7`, saved a native copy, and reopened it successfully. Focused
reverse-read recovered all three requested pagination values and an identical
restaging reproduced the native artifact byte-for-byte. This retires one more
raw-ID example and transfers one focused mutation capability, but removes no
manifest edge: all 16 ordered debts remain. Durable patches, atomic save,
aggregate peak-memory policy, fuzz/sanitizer completion, the remaining Pages
editor/example/test inventory, and complete host deletion remain open gates.

## 2026-08-08 focused Keynote slide-transition ownership

`litchi-keynote::Package` now owns selector-first read, set, native-none clear,
exact patch application, and inverse replay for existing modern slide
transitions. The public boundary exchanges complete archive-free
`transition::Settings` values and semantic slide selectors; native object IDs,
component names, protobuf values, wire records, and exact patch artifacts stay
private. A strict bounded preflight and private Buffa lazy view project the
known fields, while the accepted raw source records remain authoritative for
preservation and mutation.

Changed edits patch the modern transition leaves, preserve unknown nested
records and IWA headers, validate and synchronize the slide-node
`hasTransition` cache, reassemble only the one or two actual owner components,
and fully reopen and reverse-read the candidate under retained limits before
publication. Exact no-ops retain the source allocation, inverse replay restores
the exact source bytes, and changed legacy nested packages are refused. The
legacy host transition writer remains available for compatibility and was
hardened to maintain the same cache invariant; no manifest edge or ordered debt
is removed, so all 16 debts remain.

Apple Keynote 14.4 opened both the Rust-authored Magic Move and native-none
artifacts without repair or conversion. The inspector showed the requested
effect/timing state, native Save As/close/reopen retained it, and focused
restaging of both native-resaved artifacts was byte-identical. Public inverses
for both pristine Rust candidates restored the exact app-authored Dissolve
source. This advances a complete focused mutation vertical but does not satisfy
durable patch serialization, atomic filesystem save, the aggregate peak-memory
policy, fuzz/sanitizer gates, migration of the remaining Keynote editor surface,
or deletion of `litchi-iwa`.

## 2026-08-08 focused Pages section-text ownership

`litchi-pages` now owns selector-first read and, for rooted exact sources with
one unambiguous native body storage, whole-value set/clear, checked UTF-16 span
replacement, exact-source patch application, and inverse replay for text owned
by an existing Pages body section. The supported API retains only a
semantic section position plus archive-free text values and spans. Native body
storage IDs, section-table references, component names, package entries,
protobuf messages, raw wire records, and exact authorization artifacts remain
private. Whole-body editing is intentionally a single-section convenience;
multi-section callers select the section whose text they mean to change.

The mutation core has moved down to `litchi-iwa-text-wire`, where the focused
Pages owner and migration host share one bounded raw-storage splice without
depending on each other. The kernel preserves unknown and untouched raw
records, adjusts the complete recognized positional-table family, and reports
removed-reference provenance. Pages refuses a splice that consumes section,
footnote, or inline-object structure; graph deletion is not smuggled into a
plain-text API. A private Buffa lazy projection validates the document/body and
section-boundary graph after strict raw preflight, while raw source bytes remain
the preservation authority. Rooted exact sources with one unambiguous native
body storage rewrite one body component, fully reopen under retained limits,
and verify section text, neighboring sections, object count, and
root/section-reference topology before publication. No-ops share
the source allocation, including on legacy packages, and changed legacy nested
packages fail closed.
Changed no-root/fallback bodies also fail closed until their physical ownership
has an explicit preservation-safe mutation boundary.

The migration-host section-text methods remain available as compatibility
surfaces while their raw-ID callers, dependent-content cleanup behavior, and
legacy normalization cases are migrated deliberately. Headers, footers,
floating text, text boxes, section creation/deletion, footnote/attachment graph
mutation, and the other Pages editors also remain host work. The new focused
example and root-facade smoke move ordinary callers to semantic selectors and
typed spans, but no manifest edge is removed: all 16 ordered debts remain.

Rust integration and strict scoped Clippy evidence covers the transaction,
shared rewrite kernel, public example, root exports, exact no-op, and inverse
paths. Pages 14.4 also opened the Rust-authored emoji-bearing output without a
repair warning, saved it as a new native artifact, closed and reopened that
exact path, and rendered the complete requested text. Rust then recovered the
same semantic value, while a focused no-op and inverse over the Pages-resaved
artifact were byte-identical to it. Durable patch serialization, atomic file
publication, aggregate peak-memory policy, fuzz/sanitizer completion, an
app-authored multi-section/boundary-shift gate, native clear/range and rich
dependent-content gates, the remaining examples/tests, and the complete host
deletion gate remain open.

## 2026-08-08 amendment: cache-state transfer and focused clear/range evidence

The preceding claim that all 16 debts and the native clear/range gates remain
open is historical and is superseded by this amendment. Cache-backed
`PackageState` has transferred from `litchi-iwa` to the physical
`litchi-iwa-archive` owner. Archive ownership is bounded physical
parsed-component state; the dependency-free `litchi-iwa-cache` leaf remains
free of archive and format policy, while the host retains format/error policy.
The direct `litchi-iwa -> litchi-iwa-cache` debt identity 003 is retired
without renumbering. The current boundary count is 63 packages, 223 internal
declarations, and 15 ordered debts.

Numbers changes only one focused read boundary. `TableInfo.tableModel` uses a
strict small private Buffa projection instead of eager Prost reads
with bounded raw preflight and a required nonzero reference. Buffa does not
encode, retain unknown content, or store repeated fields, and raw source stays
authoritative. This is explicitly not a wider table-model or whole Numbers
graph migration.

Pages 14.4 opened the Rust-authored
`/private/tmp/litchi-pages-example.KdlErn/clear.pages`
(`63c2aa20f6064b9a8c5a536475d1a71b34175f4c6924a4d384f24c39fd5155e6`)
and `range.pages`
(`dd0405249a56e3e2b535e6a9541f02feda6299ce1a0959f4d68f7e44a0ae307a`)
without repair. The clear artifact was visibly empty; the range artifact
displayed exactly `Range prefix: Litchi native Pages fixture`, `Buffa lazy-view
migration verification`, and `2026-08-07`. Native Save As, close, and reopen
yielded `clear-native-resaved-20260808.pages`
(`3ba278e1934688c653ab73f1ee2a194f670545dd160aa5d8e33c2054463a9676`)
and `range-native-resaved-20260808.pages`
(`74072d9d813282618db8e47f7ebc26cc59f7c17b1abf9d22c5bbf5473b942a9f`).
Focused semantic reread matched each expected result; focused no-op and inverse
outputs over each native-resaved artifact were byte-identical to the
corresponding hash.

This advances the focused Pages clear/range evidence only. App-authored
multi-section/boundary-shift and rich dependent-content gates, durable patch
serialization, atomic publication, aggregate peak-memory policy,
fuzz/sanitizer completion, remaining ownership, and complete host deletion
remain open.

## 2026-08-08 amendment: Numbers TableInfo model-reference projection

The focused Numbers owner no longer eagerly Prost-decodes
`TST.TableInfoArchive` merely to reach `tableModel`. A two-message private
Buffa lazy projection exposes only a typed nonzero model reference. A strict
raw preflight precedes Buffa, requiring unique canonical length-delimited
`TableInfoArchive.super` and `tableModel` fields and a unique canonical,
nonzero nested `TSP.Reference.identifier`. The base drawable envelope and all
unselected TableInfo/reference metadata remain opaque caller-owned source
bytes; neither Buffa unknown retention nor encoding participates in
preservation.

The derived schema is provenance-checked against `TSTArchives.proto` and
`TSPMessages.proto`, is capped at 1 KiB of source and five generated files / 64
KiB, and fails if Buffa generates a repeated view. Its explicit bytes, field,
work, and two-level recursion budgets bound the strict and deferred scans.
Both the rooted table reader and formula-name enrichment now use the same
generated-type-free codec. Rooted failures map to the existing content-free
Numbers semantic location, while formula discovery deliberately remains
best-effort. Focused regressions cover Prost parity, opaque native metadata,
required and duplicate fields, wrong wire types, noncanonical framing, zero and
malformed identifiers, exact limits, and the checked-in native Numbers
fixture's rooted and compatibility readers. A formula-bearing constructed
package also proves that valid references still enrich sheet/table names while
malformed TableInfo metadata remains best-effort and falls back safely.

This is only a TableInfo reference seam. Table-model, tile, sidecar, and
formula payloads still use their existing bounded eager Prost paths, so it does
not claim whole-graph Buffa laziness or advance the remaining monolith deletion
gates.

## 2026-08-08 amendment: focused Keynote existing-notes vertical

`litchi-keynote` now owns semantic reads and exact-source set, clear,
insert/delete/replace, reversible patch application, and inverse replay for
text in an existing speaker-notes graph. Selection is by exact navigator name
or checked semantic position; ranges are checked UTF-16 values. Supported
callers never handle a slide, note, or storage object identifier, component
name, protobuf message, or raw record.

A private strict Buffa projection covers only the selected ownership
references. Its accepted lazy values follow bounded schema-directed raw
preflight; original records and exact IWA headers remain the preservation and
rewrite authority. Package-wide scans prove unique ownership and reject
aliases, duplicate metadata occurrences, dependent or unknown note shapes,
reserved markers, malformed selected framing, and noncanonical outer object
prefixes. A changed transaction rewrites one component and publishes only
after complete retained-limit reopening plus semantic and topology readback.
Exact no-ops retain the source, and exact inverses restore it.

The public example's set, range, and clear modes passed Apple Keynote 14.4
open, native Save As, close, and exact-path reopen without repair or conversion.
All three native-resaved packages reverse-read correctly; no-op restaging was
byte-identical, and a one-component temporary edit inverted to the exact native
hash. The focused codec and transaction suites cover strict framing, ownership
ambiguity, UTF-16 boundaries, unknown/header/ZIP preservation, limits,
conflicts, no-op replay, and exact inverse behavior.

This transfers one complete existing-graph text vertical and removes the
host's raw-ID notes example, but it does not create or delete notes graphs and
does not retire a host dependency. Boundary cleanup removes two unrelated
test-only ZIP declarations and makes every exclusively development-only
internal edge explicit; the current ledger is 63 packages, 221 internal
declarations, and 15 ordered debts. Remaining host APIs and examples, legacy
normalization, durable patch serialization, atomic publication, aggregate
peak-memory policy, fuzz/sanitizer completion, and full host deletion remain
exit blockers.

## 2026-08-08 amendment: structured-seam exit prerequisite

The attempted next deletion slice stopped at the evidence boundary. The
focused and neutral owners now have exact retained-text accounting and
semantic-boundary regressions: Pages excludes scratch slots and rendered-only
separators while reporting actual observations, Keynote charges its show title
and owned unknown animation identifiers, and empty/null topology behavior is
locked without publishing partial semantic state. The root preserves the
focused Pages observation unchanged.

The `litchi-iwa -> litchi-iwa-structured` migration edge is not retired by
this slice. Before debt 011 can be deleted, focused Numbers and root tests must
own the five surviving compatibility oracles for detached models, type-9
numeric values, global object ordering, canonical type-6001 precedence over
legacy type-6000 with deduplication, and inclusive/exceeded table limits. The
ledger therefore remains 63 packages, 221 internal declarations, and 15
ordered debts.

The three native fixtures passed a locked, read-only Apple iWork render gate
with exact post-close hashes, and one 60-second root ASan/libFuzzer campaign
completed without a finding. Those results do not replace the remaining
Numbers oracle transfer, focused deep fuzzing, aggregate peak-memory work,
edit/save compatibility, full Buffa graph migration, or the final host
deletion gate.

## 2026-08-08 amendment: debt 011 structured-read seam deleted

Deletion gate 3 advances for the structured read seam. The five blocking
Numbers compatibility oracles now live in focused and root tests, backed by a
deterministic checked-in 535-byte fixture with SHA-256
`352ca6ad6891c7222f76cdb5fe48178f1efb340dc82ab5bc6755b71a2d2595bc`.
They preserve the historically important detached-model, decimal128 type-9,
package-global ordering, canonical-then-legacy deduplication, and inclusive
table-budget behavior without depending on a host adapter.

The root facade now obtains semantic data from `litchi-pages`,
`litchi-keynote`, and `litchi-numbers`, then constructs the neutral aggregate
directly. The obsolete host module, public re-export,
`Document::extract_structured_data`, support hooks, tests, and
`litchi-iwa-structured` manifest edge are deleted. This is an intentional
breaking removal of an unpublished workspace host API at version 0.0.1; the
workspace and registry-consumer audit found no consumer requiring a temporary
alias. The neutral `litchi-iwa-structured` crate remains in its intended role.

Legacy type-6000 model admission is now strict once a bounded fingerprint
classifies a payload as model-shaped, rooted admission is budgeted before
decode with fallible allocation, and common resource failures cross the
Numbers public boundary through a content-free format-owned taxonomy. The
public-API policy rejects the retired symbols, host-index types, visible
aliases, and public glob re-exports. The current ledger is 63 packages, 220
internal declarations, and 14 ordered debts.

A locked Numbers 14.4 artifact with SHA-256
`781181e89c655da5c92b677b9ba5c939c85379e7b33ccf10e3846fe8588f9c5b`
passed a no-warning, no-save, exact-hash Computer Use read gate and confirmed
the visible sheet/table order used by the focused oracle. Focused, root, host,
policy, strict documentation, and sanitizer-target build gates form the
cutover evidence; the synthetic fixture remains authoritative for native tags
that the UI cannot expose.

This does not delete the monolith. Remaining host editors and compatibility
surfaces, focused eager Prost payloads, whole-graph Buffa lazy views, durable
patches, atomic file publication, native save compatibility, deep fuzzing, and
performance gates remain open. A 32 MiB unrelated root sidecar currently adds
approximately 32 MiB of transient RSS during prepared-source construction, so
this amendment explicitly makes no aggregate peak-memory completion claim.

## 2026-08-09 amendment: existing Keynote title/body vertical

Deletion gate 3 advances for one more Keynote mutation family. The concrete
format owner now reads and edits text in an existing slide's existing semantic
title or body placeholder through `SlideSelector`,
`slide::placeholder::Kind`, checked
UTF-16 spans, and an exact-source reversible patch. Native slide, placeholder,
and storage identifiers, component names, protobuf messages, and authorization
records remain private.

The ownership proof uses strict private Buffa lazy projections only for the
format-ownership edges. The existing speaker-notes projection supplies the
optional `KN.SlideArchive` field-5 title and field-6 body references, and the
new placeholder projection follows the required placeholder/shape inheritance
envelopes to optional `ShapeInfoArchive.owned_storage` field 4 and the
placeholder kind. The selected read forces the slide view. Package-wide proof
raw-scans every slide and note candidate. A slide candidate is forced through
the slide view only when its raw edge references the selected placeholder; the
alias scan does not force the Buffa `NoteArchive` view. Placeholder candidates
are raw-scanned and only a storage-relevant owner is forced through the
placeholder view. The scanner also rejects deprecated-storage, text-flow,
standalone shape-info, and embedded-reference aliases. Text storage decoding
and rewriting remain in `litchi-iwa-text-wire`, so this is not a whole-graph
Buffa conversion.

A changed edit commit produces output with a targeted raw-wire text splice and
a bounded invalidation of the selected `KN.SlideNodeArchive` preview cache.
The selected storage and slide node may share a component or occupy two, so
diagnostics report one or two touched IWA components. The invalidation removes
the node's rendered thumbnail fields and references, marks it dirty, and clears
only preview-owned aggregate and field data-reference occurrences in the
selected message metadata. Proven unrelated references remain exact, while
ambiguous aggregate-only ownership fails closed. The
archive owner's new bounded, exact-name deletion-aware reassembly path also
removes any root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg`; those ZIP deletions are not counted
as IWA components. The text, node, and ZIP mutations publish atomically as one
candidate. Without these invalidations, native Keynote and package
preview consumers may continue presenting a rendering made before the text
change.

This deliberately narrows the preservation claim: all other IWA objects and
retained ZIP entries remain exact, but the selected storage, selected slide
node cache records, and root previews are changed or removed by design. A
changed candidate publishes only after full retained-limit reopen, selected
semantic readback, cache invalidation and preview-absence checks,
unchanged-object comparison, and unselected-slide semantic comparison. Slides
with the separate cached title/body strings in `KN.SlideArchive` fields 37 or
38 fail closed because those fields are not yet mutation-owned. Applying a
changed patch reopens and verifies the exact target bytes already stored in the
patch; it does not reassemble them and reports the originating edit's component
count. An exact no-op preserves every cache and preview byte, shares the source
allocation, reports zero components, and deliberately skips whole-source
validation and candidate reparse.
Changed inverse application restores and verifies the complete original
artifact, including its former preview/cache state.

The obsolete host methods `set_slide_title`, `replace_slide_title`,
`clear_slide_title`, `set_slide_body`, `replace_slide_body`,
`clear_slide_body`, `set_slide_notes`, `replace_slide_notes`, and
`clear_slide_notes` are removed with their private storage-resolution helpers.
The notes removal relies on the previously accepted existing-notes vertical;
it does not claim notes graph creation or deletion. Host creation behavior,
placeholder visibility and layout, arbitrary text boxes, generic text-storage
editing, and the remaining Keynote graph editors stay in the migration host.
The removal is intentionally breaking rather than shimmed: callers move from
mutable raw-index methods to semantic selectors, checked UTF-16 spans, and
immutable `SlideTextEdit` or `SlideNotesEdit` commit flows. Inputs with shared,
ambiguous, or contradictory ownership can therefore be rejected even if the
old generic storage editor could address them.

The cache-invalidating sequential output passed Apple Keynote open, native Save
As, close, and reopen without repair, conversion, or warning. The requested
Unicode title and body plus untouched date rendered exactly, all three root
previews were regenerated, focused reread matched, and same-value title/body
transactions over the native copy were byte-identical no-ops. The Rust and
native SHA-256 values are respectively
`f3b13cd5bd614d93493cc6780ff177e6a203d990d15b9d5c592687ef40a48263`
and `cb3f9b05613505bb422942ca43e237a731454f58753ee65f26ae639187b96a6c`;
ADR 0008 records the full inverse and Computer Use gate.

This is a vertical API retirement, not a manifest-edge retirement. Title/body
placeholder creation or deletion, arbitrary text-box ownership, durable patch
serialization, atomic filesystem publication, whole-Keynote Buffa conversion,
deep fuzz completion, and complete `litchi-iwa` deletion remain exit gates.
The current metadata/policy inventory is 64 packages, 235 internal
declarations, and 14 ordered debts.

## 2026-08-10 amendment: existing Numbers table-lock mutation vertical

Deletion gate 3 advances for the focused mutation of one existing attached
Numbers table's interactive lock state. The concrete format owner now exposes
`Package::{table_lock, edit_table_lock, apply_table_lock}` over semantic sheet
and table selectors plus the archive-free `table::lock::State`. Its edit, commit,
reversible patch, diagnostics, errors, and limits keep native identities,
component names, messages, and wire values private.

The format adapter resolves semantic positions through the rooted native
document and sheet drawable order, accepting exactly one canonical type-6000
or legacy type-6003 table-info owner. The focused private codec strictly
preflights the required drawable envelope, optional canonical field-5 lock
Boolean with presence, and required nonzero model reference under byte, field,
work, and nesting ceilings. Buffa's borrowed lazy views are forced for both the
drawable `super.locked` value and table-model reference; their complete
presence-preserving snapshot must equal preflight. Raw records retain all
unknown-content and rewrite authority. This is not a Buffa migration of table
models, tiles, data lists, formulas, or the wider Numbers graph.

This supersedes the 2026-08-08 two-message, opaque-super, five-file/64 KiB
projection record. The current three-message TableInfo/Drawable/Reference
closure forces both lock and model lazy values and generates five files
totaling 83,529 bytes under an 84 KiB cap.

A semantic no-op keeps an absent lock absent and an explicit false explicit,
shares the source allocation, and performs no reassembly or candidate reopen.
A changed edit raw-patches only the selected nested scalar, rewrites one IWA
component, reassembles the exact flat package under retained bounds, and
reopens the complete Numbers snapshot before selected-state readback. Retained
fields, messages, unselected object-header metadata, components, and ZIP
members remain preservation-owned. Competing rooted sheet ownership,
contradictory selected-owner metadata, noncanonical outer object-length
prefixes, and selected merge/diff metadata fail closed instead of being
normalized. Detached/unrooted pseudo-sheet and view-state dependent references
are not owners for this rooted traversal and remain opaque and preserved.
Exact-source patches retain complete before/after
artifacts; changed application reopens the stored target, and inverse
application restores the exact original bytes. Legacy nested packages admit
reads and exact no-ops but fail closed for changed publication.

The complete Numbers-specific host read/mutation seam is deleted instead of
shimmed: direct `table_lock_state`/`set_table_lock_state`, private
`table_lock_context`, `NumbersTableInfo.lock_state` and its field-population
branch inside `tables()`, both model-specific shared helpers, and the
Numbers-only model-ID matching branch.
All Numbers state readback moves to `Package::table_lock`. The boundary checker
ratchets five exact function names with a three-under-Numbers plus
two-under-shared-helper scope and separately rejects the retired
`NumbersTableInfo.lock_state` field; the field-population and matching-branch removals
are additionally locked by compilation and compatibility coverage.
Pages and Keynote still use the generic shared getter/setter and wire codec,
and the rest of Numbers graph mutation remains deletion work. No manifest edge
or ordered debt is retired by this vertical.

The new focused example performs semantic lock/unlock selection, no-clobber
temporary-file publication, and optional exact inverse emission. The former
cross-iWork example uses the host only to construct the initial Numbers table,
then routes both Numbers mutation and readback through the focused owner;
Pages and Keynote remain host-owned in that example.

Two semantic-state tests, nine strict-codec tests, and 15 exact-source
transaction tests are present in the focused source. They inventory selector,
presence, preservation, inverse/conflict, legacy, resource, failure-atomic,
checked-native-fixture, rooted `FormBasedSheet` field path `[1, 2]`, and
concurrent-read coverage. The focused transaction suite passed 15/15,
including changed flat legacy type-6003 TableInfo publication with exact
inverse and partial-sink write accounting.
The bounded `numbers_table_lock` fuzz target compiles, and all 57 boundary
policy regressions pass. The full policy command still reports the 14
pre-existing soapberry-zip/xml-minifier annotations. A Numbers-only fuzz
package and a sustained sanitizer campaign remain exit gates.

The current writer also passed the Apple Numbers 14.4 (7043.0.93) gate. The
source SHA-256 is
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
the Rust locked output is
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`,
and inverse application restored the exact source. Numbers opened the locked
output without warning, showed `Table 1` locked with disabled cells, retained
the B2 text and B3 value 42, then completed native Save As, close, and reopen.
The native-resaved SHA-256 is
`8aa87a3afcb145b66c5c6f4e10645cd1cf658f4b65f0976612ac6d62d4652995`;
focused reread remained locked and an equal-state transaction was a byte-exact
no-op at that same hash.

This closes the focused native compatibility gate, not the exit plan's
resource and publication gates. There is no aggregate peak-memory or total
transaction-work policy covering both retained patch artifacts, rewrite
buffers, package hashing, reassembly, and full candidate reopen. A complete
transitive fallible-allocation proof remains open. The package can write exact
bytes with exact partial-sink failure accounting, and the example demonstrates
sibling-temporary no-clobber publication,
but the library does not yet own atomic durable filesystem save/replacement.
Durable patch serialization, deeper fuzzing, remaining Numbers graph
ownership, and final `litchi-iwa` deletion remain exit gates. The process-local
patch also lacks a versioned semantic operation envelope, read/write sets,
composition, three-way merge, and bounded history.
Resource/allocation errors do not yet carry the selected semantic table path,
and exact source bytes remain ordinary `Package` surface instead of an
explicit advanced/raw boundary.
The flattened `TableLock*` transaction names remain migration debt against the
focused-module short-name rule.
The archive-free `Table` snapshot does not yet carry lock state, remaining host
table/cell mutations do not enforce that state by default, and the private
Numbers locator has not converged on the neutral IWA index owner.

## 2026-08-10 amendment: existing Pages page-layout vertical

Deletion gate 3 advances for the Pages document-wide page-layout read and
mutation family. `litchi-pages::Package` now exposes `page_layout`,
`edit_page_layout`, and `apply_page_layout` over the existing archive-free,
presence-preserving `Layout`; its edit, commit, reversible patch, diagnostics,
errors, and limits keep native identities, components, message types, and wire
fields private.

The focused adapter resolves exactly one object 1/type-10000
`TP.DocumentArchive`. It strictly preflights required opaque field 15 and
layout fields 30 through 39 and 42, then forces every corresponding scalar on
the private document-body Buffa lazy view and cross-checks the entire semantic
result. The reused projection has no production encoder or repeated view,
leaves `super` opaque, and measures 122,114 generated bytes across five files
under a 124 KiB cap. Raw input retains all preservation and rewrite authority.

A changed edit patches only the layout scalars and follows the rooted raw cache
graph from required document `super` field 15 through shared-document
`view_state` field 5, the unique referenced type-210 bridge's field 1, and the
unique referenced type-10147 view-state root. Deprecated document fields 11
and 12 are rejected. The two followed local references must each have exactly
one aggregate metadata occurrence and, when present, unique field metadata at
paths `[15, 5]` and `[1]`. The invalidation removes the rooted layout-state field 1
plus its exactly owned aggregate and optional path-`[1]` reference metadata,
but preserves UI-state field 2, unrelated metadata, unknown fields, the
intermediate bridge, the detached opaque layout-state object, and detached or
unrooted view-state candidates. Missing, ambiguous, or contradictory rooted
objects or metadata, a layout/UI alias, selected merge/diff records, and
noncanonical object lengths fail closed. The selected document and rooted
view-state root may share one IWA component or occupy two, producing the
corresponding touched-component diagnostic.

The atomic candidate also deletes any root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg`, outside the component count, so
Pages cannot retain previews rendered for the former geometry. All other
retained package records and IWA content remain exact. Full retained-limit
reopen verifies layout, invalidation, preview absence, stable statistics, and
unchanged section semantics before publication. A semantic no-op leaves field
presence, caches, previews, and source bytes exact, shares the source
allocation, reports zero components, and skips reassembly/reopen. Changed
patch application reopens the stored exact target, and inverse application
restores the complete source artifact. Changed legacy nested publication is
refused while reads and exact no-ops remain supported.
Canonical unknown protobuf groups are also readable and retained by exact
no-ops, but a changed page-layout splice currently refuses a group-bearing
document payload.

The host's eager-Prost `PagesEditor::page_layout` and `set_page_layout`, its
private page-layout module/source, duplicate tests, and old example are deleted
rather than shimmed. Callers move from a mutable host editor to a focused
package edit and must chain later changes from the returned commit package.
The new example owns validated layout changes, no-clobber sibling-temporary
publication, and optional inverse output. Boundary policy ratchets the two
retired methods and module/source and forbids physical vocabulary from the
focused facade. This retires one host vertical, not the Pages editor, a
manifest edge, or the monolith.
The current inventory remains 64 workspace packages, 235 internal dependency
declarations, and 14 ordered migration debts.

The deterministic deletion gate passes all 92 Pages tests/doctests, including
10/10 focused transaction cases, plus 6/6 private codec cases, the Pages
package check, focused warnings-denied Clippy, and 63 boundary-policy tests.
The focused fuzz target compiles and completed 32 generated smoke inputs plus a
fixed changed corpus. Sanitizer execution is still an exit gate because the
installed stable toolchain rejects cargo-fuzz's sanitizer flags and nightly is
unavailable. The checked-in native Pages fixture produced a two-component,
three-preview-deletion 792 by 612 point landscape candidate while preserving
semantic text; inverse application restored the exact source. Source and
candidate SHA-256 values are
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`
and `79e00545ef6e2e30e366e3160b7d9126bf06cffac5fbbd5551e3d3789cc298e4`.
Apple Pages 14.4 (7043.0.93) opened it without warning, showed US Letter landscape,
Document Body, and all three exact fixture lines, then completed native Save
As, close, and reopen. Native Save As regenerated all three previews and
produced SHA-256
`8228e7518bb080bd8e5ec134d0abc7484c8825ad3cde3d16cabf76c5dbd8ef82`;
focused equal-layout readback was a byte-exact no-op with zero components and
zero preview deletions.

The opaque detached layout-state object, other Pages settings and render
caches, durable/mergeable patches, whole-graph Buffa coverage, and final host
deletion remain exit work. So do aggregate transaction peak-memory and total-
work accounting, a complete transitive fallible-allocation audit, and a
library-owned atomic durable filesystem replacement. The example's synced
temporary/no-clobber workflow does not close that library contract. Exact
source-byte exposure and the flattened `PageLayout*` public names also remain
API-boundary debt.

## 2026-08-10 amendment: combined Pages document-settings vertical

Deletion gate 3 now also passes for the combined Pages document-options and
footnote-settings family. `litchi-pages::document_settings::Settings` is an
archive-free composite of `document_options::Options` and
`footnote::Settings`; its focused module owns canonical short `Edit`, `Commit`,
`Patch`, `Diagnostics`, `Error`, and `LimitKind` names. The only package entry
points are `document_settings`, `edit_document_settings`, and
`apply_document_settings`; their focused method/type signatures expose no
native identifier, component, source-byte, archive/IWA, core/proto/Prost,
Buffa, or generated type.

The private seam resolves the unique document root and its nonzero local
`TP.DocumentArchive.settings` field-7 reference to the unique type-10012
`TP.SettingsArchive`. That edge must occur exactly once in aggregate metadata;
optional field evidence must be unique and match path `[7]`. Strict raw
preflight is cross-checked against forced Buffa lazy views for fields
1/2/3/9/10/30-34: body, headers, footers, hyphenation, ligatures, footnote
kind/format/numbering/gap, and facing pages. Five generated files measure
174,682 bytes under the 176 KiB cap, with deterministic aggregate SHA-256
`7618a60db84b87e28eea67a8acd85ce8eb19513cf4cee7654c1c4e78f405f824`;
there is no repeated view or production encoder.

A semantic no-op is byte-exact, shares the source, reports zero touched
components and deleted previews, and performs no reassembly or reopen. A
changed edit patches the settings owner, invalidates the exact rooted
document/view-state cache chain, and atomically removes root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg`. Depending on component placement it
rewrites one or two components, with preview deletions diagnosed separately.
The retained-limit reopen verifies settings, invalidation, preview absence,
stable statistics, and preserved semantics. Changed patch application reopens
the stored target, conflicts fail, and inverse application restores the exact
source.

Canonical unknown scalars are retained. Bounded canonical groups can be read
and survive exact no-ops, while changed group-bearing splices are deliberately
refused. Noncanonical and wrong-wire encodings, duplicates, invalid scalar or
reference encodings, contradictory selected-owner metadata, merge/diff state,
and malformed object framing fail closed. Legacy nested `Index.zip` sources
remain readable and exact on no-op, but a changed transaction returns
`UnsupportedSource`. That explicit policy supersedes the deleted host path's
changed normalization behavior.

The deletion removed `PagesEditor::document_options`,
`set_document_options`, `footnote_settings`, and `set_footnote_settings`;
`document_options.rs`, its nested `document_options/wire.rs`, and
`footnote_settings.rs`; and two host examples plus duplicate tests. A single
focused example now demonstrates immutable chaining, no-clobber sibling-temp
publication, and optional exact inverse output. The boundary retirement/public
leak ratchet passes 70/70 tests; the live repository checker retains only 14
unrelated pre-existing diagnostics (12 for six `soapberry-zip` dev edges and
two for `xml-minifier`). This retires the combined vertical, not the remaining
Pages editor, a manifest edge, or the IWA monolith. Inventory remains 64
workspace packages, 235 internal dependency declarations, and 14 ordered
migration debts.

The deterministic gate is green: 108/108 Pages tests/doctests, 14/14 focused
transactions, 4/4 codec tests, 6/6 facade tests, package check and docs, and
no-dependencies warnings-denied Clippy. The fuzz target compiles and passes
no-op and changed smoke runs; sanitizer execution remains blocked by the
stable-only toolchain rejecting the required flags with no nightly installed.

The native gate used Apple Pages 14.4 (7043.0.93) and a fresh app-authored,
footnote-bearing seed, SHA-256
`9da01e2805459e05450551827140069eefe8049aeeacc7625d3c62d7e00ffeab`.
The Rust candidate, SHA-256
`3d052e7f1ec86e57ea0553e46f628de1d9fa5bdda615ded9410fca29c93f0995`,
reported two touched components and three deleted previews; inverse restored
the exact seed. Pages opened without warnings and confirmed body/header/footer
and facing pages enabled, hyphenation and ligatures disabled, Roman footnotes
restarting each page at an 18-point gap, and the three body markers plus note
unchanged. Native Save As, close, and reopen preserved those values,
regenerated all previews, and produced SHA-256
`803167e2479c459f9a33c8ecfc4d713f596fdc5d5d337090ab3c90e467a0cba6`.
Focused same-settings readback was byte-exact with zero touched components or
deletions, as was its inverse.

Exit work still includes shared aggregate transaction peak-memory/total-work
accounting, the infallible retained `ArchiveInfo` clone, a complete
fallible-allocation proof, group-aware changed splicing, exact streaming and
partial-output accounting plus a robust Pages `Package::write_to`, and a
library-owned atomic durable filesystem replace. Patches still need versioned
deterministic serialization, semantic operations and read/write sets,
composition, merge, and history. Exact source bytes remain ordinary `Package`
surface, and opaque cache objects plus remaining Pages settings/render state
remain exit work.

## 2026-08-10 amendment: hardened Keynote show-settings deletion gate

The earlier partial Show gate is superseded. `litchi-keynote::show` now owns
the complete archive-free focused family `Settings`, `Edit`, `Patch`, `Commit`,
`Diagnostics`, `Error`, and `LimitKind`. The package's `show_settings`,
`edit_show_settings`, and `apply_show_settings` signatures expose no raw
source, native identity, IWA member, or generated type. `Edit::set` consumes
the edit for immutable chaining, and callers emit a returned package through
bounded `Package::write_to` rather than obtaining raw source bytes.

The private ownership chain is the unique root `Document.iwa`/object 1/type-1
`KN.DocumentArchive`, its required local field-2 show reference, then one
referenced object in exactly one component with one type-2 `KN.ShowArchive`.
The nonzero reference must occur once in aggregate metadata; any field-local
evidence must be unique at `[2]` and cannot compete on another path. External,
missing, duplicated, or contradictory selected ownership fails closed. A null
root show reads as default settings but cannot be materialized by this edit, so
only its exact no-op is supported.

The root and show readers each run strict raw preflight before forcing private
Buffa lazy views and cross-checking the full selected values. Root provenance
is five generated files/58,630 bytes under 60 KiB, aggregate SHA-256
`7918aad2578cf3bd07eb0be36f2e31d11f93391584308c1e4adc1fd86ed065fd`;
show provenance is five files/138,661 bytes under 140 KiB, SHA-256
`747fe9f99dc5bb1855aae1bfcb16065a5fe6305bdbf8730a21ef24bb75e915ee`.
The complete known Show/SlideTree envelope and slide limit are validated, but
the repeated slide tree is hand-routed and never retained by generated code.
Ratchets forbid repeated views and production encoders. Exact raw records own
preservation.

A changed publication additionally requires canonical object framing and
rejects selected merge/base/diff state. It raw-splices only the presentation
size and eight optional show scalars, rewrites one IWA component, and fully
reopens/verifies the retained-limit candidate. Size and slide-number changes
remove any root `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`;
playback-only changes preserve them. All slide components and slide-node
thumbnail/playback caches remain exact under either policy. Component and
preview counts are diagnosed independently.

Semantic no-ops preserve every byte/cache, share the source, and skip
reassembly/reopen. Changed patch application verifies exact source and stored
target state before reopening the target; inverse application restores the
exact complete source. Legacy nested `Index.zip` retains reads and exact
no-ops, but changed edits intentionally fail with
`show::Error::UnsupportedSource` under Preserve policy instead of running the
old normalizing writer.

Deletion gate 3 removes `KeynoteEditor::show_settings` and
`set_show_settings`, the editor `show_settings` module and source,
`examples/edit_keynote_show.rs`, and their direct mutation/compatibility tests.
The focused example now owns semantic staging, exact inverse, distinct-output
and no-clobber temporary handling, and `write_to`. Boundary ratchets prevent
the host surface or physical focused-API leaks from returning.

This does not delete every host Show consumer: read-only
`KeynoteDocument::show` still decodes a Prost `KN.ShowArchive`, and other
creation/slide/media/transition/soundtrack/graph paths remain. Thus the direct
editor mutation vertical exits without retiring the monolith, a manifest
edge, or an ordered debt.

The current deterministic gate passes 19/19 focused show-settings cases,
106/106 full codec cases, 49/49 focused Keynote codec cases, Keynote all-target
checking, the host library check, umbrella facade compilation, strict rustdoc,
and 80/80 boundary regressions. Both focused live audits are empty; the general
repository checker retains only 14 unrelated pre-existing diagnostics. The
fuzz target passes `cargo check`; its stable-built executable completed 32
bounded cases with expected missing-sanitizer-symbol warnings. A sanitizer run
through cargo-fuzz remains unavailable because it requires nightly.

Apple Keynote 14.4 (7043.0.93) passed two native gates from exact source
`f3adcde9315b6df580805bcb63c995cc1e1ef569a4befa06a102485e13c883b2`.
The pristine slide-number Rust candidate
`6d28d461c1203f00384fe6a758df1f903c7555b90ff02d2dc32d856aa9056c13`
became `031a701040ed1ea9a5111fe3e298bcddcf33d498891f827b703d01328ba17224`
after native Save As/close/exact-path reopen. The pristine Custom 1280-by-720
candidate `67e9ff0557683af105dfe57f999acabcde23f121f7aebb06102c93e03121c027`
became `a3a2f6e072db4bd952f2c02e528f25c3656dba5810fbff75e93b5a699aac0eda`.
Both Rust inverses restored the exact source. Both artifacts opened without
repair/recovery/conversion, auto-played, and retained Self-Playing, Loop, Play
on Open, five-second transitions, two-second builds, and their respective
Widescreen 1920-by-1080/Custom 1280-by-720 inspectors through exact-path
reopen.

Each Rust candidate deleted all three root previews; Keynote regenerated them
on resave. All four `Index/Slide*.iwa` hashes remained exact across each
candidate/resave pair. Keynote did normalize explicit slide-number true to
absence: restaging absence is an exact `031a7010...` no-op, while restaging
true changes the artifact. Same-settings no-op and inverse on the native size
resave are exact at `a3a2f6e0...`. The native gate therefore proves slide-cache
preservation and conservative root-preview invalidation, not persistence of
the slide-number scalar.

Exit debt remains in the host Prost Show read/other generated consumers,
aggregate transaction peak-memory and total-work accounting, complete
fallible-allocation proof, group-aware changed splicing, stable versioned
semantic patch serialization with read/write sets/composition/merge/history,
and library-owned atomic durable filesystem publication. `write_to` is bounded
exact streaming, not flush/sync/rename/durability. A full sanitizer-backed fuzz
campaign remains explicit verification work.

## 2026-08-10 amendment: Numbers names mutation exit

Deletion gate 3 now passes for the public Numbers editor sheet/table rename
surface. The focused owner is nested
`litchi-numbers::names::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`;
the umbrella facade keeps `litchi::numbers::names` rather than flat aliases.
`Package::edit_names` is an infallible empty batch, consuming stages resolve
semantic sheet/table selectors against one immutable base, and
`Package::apply_names` owns exact replay/inverse. No native ID, component,
archive/generated/wire value, or raw source slice crosses these signatures.
`source_bytes` is crate-private and `write_to` owns exact output.

The mutation graph is rooted from TN document field 1 through the local
Sheet/FormBasedSheet sequence. A table traverses the rooted sheet drawable
path `[2]` or `[1, 2]` to one TableInfo, then required field 2 to one
TableModel. Each followed edge needs exact aggregate reference metadata and
optional unique matching field evidence; every selected model needs one and
only one rooted TableInfo owner. Strict raw decoding is cross-checked against
forced Buffa views for sheet/form names and TableModel identity/name. The
projection has five generated files/82,641 bytes and deterministic SHA-256
`944b7637fd6bf0eb895174b1e9229aa9eb9c393e05c666a86dd2843792eefe3e`.
Raw records remain the preservation owner.

The edit validates the final batch, so swaps and collision-away renames are
atomic while duplicate targets and final sheet/table namespace collisions fail
without publication. Changed table renames refuse selected table locks, any
rooted pivot owner, and rooted volatile sheet/table-name dependencies. A
sheet-only rename remains allowed when an unselected table is locked. The
native Θ(T²) pivot dependency traversal is conservatively work-bounded before
native scanning. Touched components are grouped and rewritten once, followed
by complete reopen and exact locality verification.

Changed batches delete every existing root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg`, diagnose previews separately from
components, and preserve `Index/ViewState.iwa` plus unrelated ZIP/IWA state
exactly. No-ops share the source and skip changed guards/cache/reassembly.
Changed patch application reopens the exact stored target; inverse restores
the source including previews. Canonical/form and accepted legacy native
message variants remain supported when rooted ownership is unambiguous.
Nested legacy physical packages retain reads/exact no-ops but changed rename
fails as `names::Error::UnsupportedSource`.

The host deletion removes `NumbersEditor::rename_sheet`, `rename_table`, their
direct tests, and `examples/rename_numbers_items.rs`. The focused example now
owns semantic batch selection, exact inverse, bounded `write_to`, and synced
no-clobber publication. The private `rename_attached_table_in_package` helper
remains for Numbers sheet duplication, and its `rename_table_in_package`
wrapper remains because Pages and Keynote attached-table flows consume it.
Therefore this exits the public Numbers editor mutation family, not every
shared native rename helper, and removes no manifest edge. Ordered debt 015
(`litchi-iwa -> litchi-numbers`) remains; inventory stays at 64 packages, 235
internal dependency declarations, and 14 ordered debts.

The deterministic gate passes 10/10 focused tests, 105/105 Numbers library
tests, the 1/1 root-facade test with `--features numbers`, 89/89 boundary
regressions, both live focused audits, `litchi-numbers --all-targets` checking,
`litchi-iwa --lib` checking, and strict rustdoc. Host
`litchi-iwa --all-targets` is not claimed because unrelated examples remain
red. Stable fuzz build plus eight bounded control-flow runs passed with
expected missing sanitizer symbols; that smoke is not ASan.

Apple Numbers 14.4 (7043.0.93) opened source
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
changed to pristine Rust candidate
`22f8bc21223317318ec23ec764b8998af77a2c7800c68cbe88351abdb26b6e56`
without warning, repair, recovery, or conversion. Public inverse restored the
source. The unlocked table was selectable/editable; UI showed sheet
`Líneas 你好 🧪`, table `表 Café №42`, exact B2 and B3=42. Save As, close, and
exact-path reopen produced
`e1803b0568454a345f7962c5b4c72e8cb3d78adb2c87d5db1e6c58288a9413c4`,
regenerated three previews, and retained the data. Equal restage, no-op apply,
and inverse were byte-exact at the resave hash.

The independent locked oracle
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`
reported `Locked`/`Locked items cannot be edited`, disabled cells, enabled
Unlock, and no title change after Edit. This is the native protection evidence
for table-rename refusal and the sheet-only exception.

Exit debt remains in the bounded but native Θ(T²) preflight, aggregate
peak-memory/total-work accounting, complete fallible-allocation proof,
process-local complete-artifact patches and absent stable semantic patch
serialization/read-write sets/composition/merge/history, library-owned atomic
durable publication, and sanitizer-backed fuzzing. `write_to` does not itself
flush, sync, rename, or make output durable.

## 2026-08-10 amendment: Keynote transition host-mutation exit

Deletion gate 3 advances for direct Keynote slide-transition editing. The
focused owner exposes canonical nested
`transition::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` plus
selector-first `Package::{slide_transition, edit_slide_transition,
apply_slide_transition}`. No raw source, native ID/component, generated type,
or wire representation crosses the focused surface; `write_to` owns exact
output.

Changed edits prove the rooted Show/SlideTree `[3, 2]` reference to the
selected SlideNode and its field-2 reference to the selected SlideArchive.
Unique object/component/type ownership, exact aggregate reference metadata,
optional unique matching path metadata, strict semantic/node-marker agreement,
canonical selected framing, and absence of merge/base/diff state are required
before publication.

The rooted audit walks the Show's slide-node list once and resolves every node
through the package's sorted, globally unique object index. This bounds lookup
cost to `O(slides log objects)` and charges aggregate node-message plus
reference-payload bytes to `LimitKind::WireWork`, without a per-node reset.

Strict raw preflight precedes a five-message private Buffa lazy-view
cross-check. Its 2,347-byte derived schema is tied to the canonical KN field
declarations, contains no repeated projection or production encoder, and
generates five files/208,052 bytes under the 224 KiB ceiling. The validated raw
records retain preservation and splice authority. A single aggregate field
budget and strict-plus-Buffa work budget cover the selected SlideArchive,
transition, attributes, and animation envelopes, so nesting does not renew
either allowance.

Only SlideArchive transition field 4 and, when effect presence changes,
SlideNode `hasTransition` field 7 may differ. Co-located owners rewrite one
component and split owners at most two, once each. Full reopen and exact
locality checks preserve all unselected objects/messages/members, unknowns,
metadata, all three root previews, `Index/ViewState.iwa`, and slide/node
playback caches. Transition changes are playback-only and do not invoke root
preview deletion. Semantic no-ops share exact source state. Clear on an
already absent transition is idempotent; changed nested legacy sources refuse
with `transition::Error::UnsupportedSource`. Exact patch apply/inverse retains
the complete artifact contract.

The deletion inventory is the three host methods `slide_transition`,
`set_slide_transition`, and `clear_slide_transition`; the
`transition_lifecycle` module/source; clear/edit/set-effect host examples; and
five whole direct lifecycle/CRUD/locality mutation tests. That exact host scope
changes by +120/-998 lines, net -878. The focused edit example replaces the
mutation workflows.

The host still owns `KeynoteSlideInfo.transition` and slide read/decode paths.
`transition_wire.rs` remains specifically for `KeynoteEditor::slides()`
aggregate decoding and no-op validation; creation uses the separate
`creation.rs::transition()` helper and retained creation example. This
therefore retires direct editor mutation, not all host transition ownership or
the monolith. No manifest edge is removed: debt 014 remains and the inventory
is unchanged at 64 packages, 235 internal declarations, 14 `litchi-iwa`
dependency declarations, and 14 ordered debts.

The exit gate passes 8/8 focused transition tests, 79/79 Keynote library tests,
6/6 warning-denied doctests, 7/7 root-facade tests with `--features keynote`,
6/6 codec tests, and retained host conversion/reader tests at 3/3 and 7/7.
Common exact-artifact/batch infrastructure passes 10/10 focused and 140/140
full tests plus strict library Clippy; archive coverage reports 79 unit and 2
integration tests. `cargo check -p litchi-keynote --all-targets`,
`cargo check -p litchi-iwa --lib`, host no-run, formatting, diff, and 101/101
boundary gates pass. Every fuzz bin checks, while the generated no-op,
fixed-clear, and fixed-set stable executables each ran six
bounded cases. Their expected missing-sanitizer-symbol warnings preclude an
ASan claim.

Apple Keynote 14.4 (7043.0.93) opened disposable copies without warning,
repair, recovery, or conversion. Source SHA-256
`ab186d8d59c858e1b3c2596fd45463cec75ddd92e9fda9032da656a940e68dca`
produced pristine Magic Move
`d5d24386cb544374f4c26da4349f7be961be34180a4536578616886a56af8c1a`
and clear
`5235a3d03dbabced6d06a03b4873826da8602d97f478c61f6467b35d732a08e5`;
both inverses restored the source exactly. Magic Move displayed 2 seconds,
Automatic, and a 2.25-second delay; clear displayed No Transition Effect while
retaining Automatic and the same delay. Save As, close, and exact-path reopen
preserved both inspector states.

The native resaves were
`dda5049cf431b5c88ea0a9fb209c67edc0d7f0764c23a17eb4e9fdf947d786f6`
and `784069ca8bd2729829bcf204cccdced93f7fbea2b5f8c6b3e4965b47ef423e94`.
Equal restaging over each reported `changed=false` and
`touched_components=0`; exact comparison, output, and no-op inverse retained
the corresponding native hash. Remaining exit debt is aggregate
peak-memory/total-work and complete fallible-allocation proof, process-local
complete-artifact patches without stable semantic serialization/read-write
sets/composition/merge/history, library-owned durable atomic publication, and
sanitizer-backed fuzzing. `write_to` is not a durability boundary.

## 2026-08-10 amendment: Numbers table-header host-mutation exit

Deletion gate 3 advances through Numbers table-header settings. The semantic unit
already exists as archive-free `table::headers::{Count, Settings}`; the focused
owner adds canonical nested
`table::headers::transaction::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind, Path, InvalidReason}` types and
`Package::{table_header_settings, edit_table_headers, apply_table_headers}`.
Read/edit selection takes an explicit sheet plus sheet-scoped table rather than
the host's workbook-wide catalog. `Edit::settings` borrows staged state;
infallible consuming `Edit::set(self, Settings) -> Self` replaces it. No physical/native vocabulary
or new source artifact accessor belongs in those signatures, and exact package
output remains `write_to`.

Changed admission must prove the rooted Document field-1 Sheet/FormBasedSheet
owner, its `[2]`/`[1, 2]` TableInfo path, and TableInfo field-2 TableModel
reference. Local edges resolve uniquely with exact aggregate and optional
matching field metadata; competing rooted TableInfo ownership and selected
metadata contradiction are refused, while detached/unrooted pseudo-owners stay
opaque and exact.

A changed transaction refuses a locked selected table, enforces present count
range `1..=5`, ensures header plus footer rows and header columns fit the
declared table dimensions, and retains absence versus explicit values for
fields 9/10/11/12/13/29/32. Strict selected framing, all finite resource
ceilings, and complete locality verification are part of the deletion gate.

The selected raw record is cross-checked through five private Buffa generated
files/51,480 bytes with no repeated views and SHA-256
`5a94caa4620c56bb464792084c01325cef01744bebac97ef948466b9dea105dd`.
Raw bytes remain authoritative.

Field-85 pivot state blocks every change. Fields 81/84/86 or nonempty 83 block
header counts; active field-81/83/86 category/group state also blocks section
counts. Strict TableInfo aliases 4/5/7/8/15/16/17 gate their corresponding
header/section counts, rooted HeaderNameMgr gates header counts, and deprecated
sheet field 4 gates repetition. Each refusal is `UnsupportedDependency`.
Footer/freeze/repeat and dependency-free counts stay in scope. For admitted
changes, only selected TableModel header fields are authorized to differ; this
does not claim that all native counts have a TableModel-only closure.

Admitted changes rewrite the selected component once and delete each existing
root preview because header settings affect rendering.
`Index/ViewState.iwa`, unrelated objects/messages/members, unknowns, and
detached state remain exact. A no-op shares the source and preserves previews;
changed apply reopens its exact retained target, while inverse restores the
complete source and previews. Changed apply first matches the retained selected
source payload and preflights conservative source-plus-target transaction work.

The host-retirement inventory is exactly the public Numbers editor header
read/write pair, two whole dedicated mutation tests, one duplicated `Count`
unit test, and `edit_numbers_table_headers.rs`. Ten mixed structural/sort tests
and seven creation/topology examples survive via private helpers or focused
package handoffs. The `table_headers` module/source, wire codec, attached
helpers, package bridge, row/column/sort callers, and Pages/Keynote owners
remain; this is not module deletion.

The focused replacement's private package code is now separated into `api`,
`dependencies`, `error`, `ownership`, `resolve`, and `rewrite` modules, each
under 600 lines. This changes neither the canonical public namespace nor the
host-retirement boundary. Category-owner group metadata is traversed once and
then resolved, giving a bounded linear declaration proof instead of an
`O(groups * references)` rescan.

Rooted canonical/legacy roles remain supported when unambiguous, while changed
nested legacy physical packages return `UnsupportedSource`. Locked reads and
no-ops remain valid; changed edits refuse and invalidate root previews. No
manifest edge is removed, so debt 015 and the current 64 packages/235 internal
declarations/14 ordered debts remain.

The native refusal oracle used Numbers 14.4 to change source
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
to two header rows/columns without warning while preserving B2/B3. Its
136,213-byte save
`5c2323b509e5ea9a975b5f254bbd46cf42657aa1c3858d2c7e98f30f07e4b40c`
changed TableModel, HeaderNameMgr, a new manager tile object, and CalcEngine
formula/dependency state. This supports fail-closed dependency refusal, not a
Rust writer or native count-parity claim.

The separate freeze oracle toggled Freeze Header Rows off, preserved 1/1
counts and B2/B3, and saved 136,199 bytes at
`015568e6b922e80fbfb760491dc49994ccc2218356ed197131beb46c1bd75850`.
Only TableModel 904538 field 12 moved from true-present to absent and the
HeaderNameMgr stayed exact. A native off-to-on control produced
`df44ed7d0b12c1d372dad7ad7361ed1140d41967921ee42b71a4072b78615721`.
Native Save regenerated equivalent ViewState with different IDs, so this is
compatibility evidence, not raw ViewState equality.

The exit gate passes 8/8 focused tests with default features and 8/8 without,
4/4 codec tests, 2/2 facade tests with `--features numbers`, and 114/114
boundary regressions. `cargo check -p litchi-numbers --all-targets`,
formatting, diff, warning-denied no-dependency rustdoc, and doctests (one compile-fail pass, one
ignored example) are green. Strict Clippy reports no new header-file finding;
unrelated baseline codec/extractor/table warnings keep the full crate gate red.

The fuzz target checks and its stable fixed-input control-flow smoke ran eight
cases with expected missing-sanitizer-symbol warnings; this is not a
fuzzing/ASan result and no nightly run occurred. Focused CLI source/inverse
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
produced candidate
`a8b88d21806b547a5265c60662610f68f524173cac1ca4252d368596c8ef8d2a`,
diagnosing changed=true, one touched component, and three deleted previews.
No native UI open of that Rust candidate is claimed.

A separate post-split freeze-row-only candidate, SHA-256
`c938d74bcf04be692097488af838f5105a8470e337eafa06fdc8b94b36231d6a`,
opened through Computer Use in Numbers 14.4 without repair or warning. Table 1
remained 22 by 7; header columns/rows/footer rows were 1/1/0; Freeze Header Rows
was unselected; and B2/B3 retained the fixture text and 42. Its inverse matched
the pristine source bytes exactly.

Exit debt remains in aggregate memory/work and complete fallible allocation,
process-local full-artifact patches without stable semantic serialization/
composition/merge/history, library-owned durable atomic publication, baseline
Clippy cleanup, and sanitizer-backed fuzzing. No dependency edge or debt item
is removed by this vertical.

## 2026-08-10 amendment: Keynote placeholder-visibility host exit

This Keynote host exit moves title/body visibility ownership to
`slide::placeholder::{Kind, State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and semantic `Package` read/edit/apply methods. The focused API
signatures do not expose generated messages or raw source artifacts. A missing
role is readable as `None` but cannot be manufactured by this transaction.
The canonical selector is now the shared
`slide::placeholder::Kind::{Title, Body}` for both visibility and slide-text
operations. Replacing `SlideTextRole` is an intentional source break; the
common discriminator still fronts distinct ownership and mutation contracts.

The vertical owns only the selected title/body reference's membership in
SlideArchive owned-drawables field 7 and z-order field 42; stable role
references remain fields 5/6. Rooted Document/Show/SlideNode ownership, exact
reference metadata, co-location, and strict raw plus Buffa placeholder
projection are required. Changed edits fail closed on aliases, conflicting
list evidence, merge/framing state, cache/layering state, layout overrides, and
selected-placeholder builds.

A change rewrites one or two components depending on SlideNode co-location,
invalidates the selected node cache, and deletes all three root previews while
preserving `Index/ViewState.iwa` and unrelated data. No-ops and inverses are
byte-exact; changed patch application validates the retained source payload.
This is not a transfer of slide-number, layout, placeholder creation, text-box,
or style mutation.
Ownership uses linear payload occurrence/kind and metadata declaration indexes;
the bounded 4,096-to-8,192-object step stays within 2.3x recorded work. A
budget-aware single SlideNode pass conditionally invalidates and exact-verifies
the direction-aware delta. Verification uses only bounded, fallibly allocated
occurrence/declaration indexes, with no full node/payload clone or verification
rewrite. Zero allowance fails atomically before publication. Structural work
charges every `MessageInfo` and `FieldInfo`, including empty records; 4,096
empty `FieldInfo` records are rejected atomically under zero and payload-only
allowances. The slide router precharges exact
`source + output + 2 * fields` work before allocation.
Full precharge covers selected/nonselected payload bytes, metadata vectors,
paths, features, and bases, every aggregate/`FieldInfo` reference in both
`Work` and `References`, and `header_length`. Low allowances atomically reject
a 256-KiB sibling plus 2,048 references/vectors.

Keynote 14.4 native oracles use pristine 500,058-byte SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Reshowing title changed
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
to
`9d914ea25a42aaced4459a429e776b09b2024e2858133369f159dad7bce67325`
and appended title after body; reshowing body changed
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
to
`8ee6ac8230273def64450b4cee86c9678849d77b5a7fbd11eb88e0c786279eee`
and appended body. UI checkboxes, canvas, date/other role, and close/reopen were
confirmed. Native cache regeneration is compatible with the focused writer's
conservative preview/cache invalidation.

The Rust-authored title-hidden candidate
`df119410433b97b9993d46619764a8ffb75f257b16c0680cd54faabd9a453cdd`
reported changed=true, two touched components, and three deleted previews, and
its inverse restored the exact pristine hash. Keynote 14.4 opened the candidate
without warning with Title off, Body on, and body/date retained. Save As, close,
and reopen preserved that state in the 475,102-byte native resave
`c5c996415191758b9fc638a8fdf024a912a6fe2ac4c3989970f0cb611e0670e3`.

The reverse-direction Rust gates are also exact. Apple-hidden title
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
became shown
`3d36d31c6222b7622cab180f6dd9559ccf43f4b481e6b245c9d2c56fe8852b2c`,
and Apple-hidden body
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
became shown
`3e8855e954c16bd32350e057665b5ee4758a02e85ad23c3c6543f1caef177b13`;
each inverse restored its exact hidden source. Both shows diagnosed
changed=true, two touched components, and three deleted previews.

The completed host exit removes
`KeynoteEditor::{set_slide_text_placeholder_visible, set_slide_title_visible,
set_slide_body_visible}`, public `KeynoteSlideTextPlaceholder`, the full 150-line
`keynote/editor/placeholder_visibility.rs` module/source, two whole direct
tests and their exclusive constant, and the 30-line
`set_keynote_placeholder_visibility` example. Five mixed layout assertions now
use focused reads. Shared placeholder ownership, layout, and slide-number code
remain, so this vertical does not overstate their retirement.

The final gate passes 94/94 Keynote library tests, 18/18 filtered slide-preview
tests, 5/5 focused visibility tests, 25/25 slide-text integration tests, 8/8
root facade tests with `--features keynote`, 7/7 doctests, and 129/129 boundary
regressions. Keynote all-target and host-library checks, warning-denied library
Clippy and rustdoc, formatting, and diff checks are green. The expanded
`keynote_slide_text` fuzz target compiles and completes a bounded stable
control-flow smoke; expected missing sanitizer symbols mean this is not
sanitizer-backed fuzz evidence. Native and exact inverse gates complete the
compatibility proof. No dependency edge or debt item is removed.

## 2026-08-11 amendment: per-slide Keynote slide-number host exit

The next completed host exit reuses the canonical
`slide::placeholder::{Kind, State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` transaction and Package read/edit/apply facade for
`Kind::SlideNumber`. This supersedes only the title/body amendment's statement
that slide-number mutation remains in `litchi-iwa`. The global Show field-6
preference remains with `show::Settings`; slide layout, placeholder creation,
slide text, and style mutation remain outside this transfer. The slide-text
owner rejects `SlideNumber`, so the shared `Kind` discriminator does not merge
the distinct operation and ownership contracts.

The format owner now proves the rooted Document field-2 -> Show/SlideTree
`[3,2]` -> SlideNode field-2 -> SlideArchive path and the selected
SlideArchive field-20 native-kind-1 placeholder. Canonical Node field 18 must
agree with exact selected membership in both Slide fields 7 and 42. A show
appends one reference after each existing field group; a hide removes only
those references. Global scanning rejects competing rooted ownership, aliases
with title/body/object/template/build/style or storage dependencies,
contradictory membership, missing placeholders, and unsupported style/storage
closures. Exact hidden no-ops preserve absent/false representation; changed
hides use canonical false and remain exactly invertible through retained patch
artifacts.

Storage zero is an allowed native closure and never becomes metadata ref zero.
A nonzero storage must remain in the selected component and satisfy the strict
type-2001 storage/type-2043 attachment proof: absent/3 storage kind,
`in_document=true`, one U+FFFC, one attachment entry at character zero, exact
aggregate metadata/dependency paths, empty attachment textual super,
absent/zero attachment kind, and no attachment object refs. Legacy nested
packages retain read and exact-no-op compatibility, while a changed mutation
returns `UnsupportedSource` rather than normalizing them.

The Buffa seam adds `KNSlideNumberArchive.proto` for Node field 18, bounded
storage scalars and borrowed attachment table, and attachment super. Strict raw
parsing precedes forced/cross-checked lazy views; no repeated generated view or
encoder owns preservation. Rooted/storage validation and scalar splice/delta
verification are split into dedicated submodules under the focused visibility
and preview owners. Generated-build evidence is five files, 112,101 bytes,
zero repeated views, a 116-KiB cap, and digest
`eacce4103b5c9f9f32fd98639b81249ae1d15fcd63da6fe636569e0a2a324c30`.

Limits charge codec bytes/fields/work/depth, rooted object/payload/metadata
scans, aggregate and field references, selected and nonselected payload bytes,
bounded fallible indexes, output allocation, exact bidirectional delta, and
archive reassembly. There is no full node/payload clone or second verification
rewrite. Allocation/limit failures are typed, content-redacted, and atomic.

A changed operation touches the Node and Slide components, deletes all three
existing root previews, reassembles, and reopens. It deliberately preserves the
Node thumbnail/cache; only field 18 and selected field-7/field-42 membership
change. ViewState, storage/attachment closure, other slides/roles, and global
Show field 6 remain exact. A no-op shares source and skips reassembly/reopen;
changed apply exact-validates its source and stored target before reopening.
The focused output is `write_to`; process-local patch serialization,
allocation/peak-memory, work-bound refinement, and durable save remain shared
debts.

The exact host deletion is one public
`KeynoteEditor::set_slide_number_visible` method, its full 172-line
`slide_number` source/module, the 23-line mutation example, and two direct
whole tests plus four constants and their fixture helper. The 53-line creation
example survives and moves only its second-slide edit through the focused
Package. `KeynoteSlideInfo` read state, creation builder/tests, shared
ownership, layout, title/body visibility, and global Show settings remain.
This removes no manifest dependency edge and closes no recorded debt item.

Native compatibility starts at 500,058-byte
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Rust's 455,859-byte shown candidate is
`a2dafcd4ffc57bafc3bbf7d7cd4ee8131bab2c06dd52adc292632d4208c126be`,
with changed=true, two touched components, three deleted previews, and an exact
inverse to the source. Keynote 14.4 (7043.0.93) opened without warnings,
displayed attachment `1`, checked Slide Number, and preserved title/body/date.
Save As, close, and exact-path reopen preserved that state in the 500,192-byte
resave
`b1edd073d309157d27508baf4aedbe93d6dee0687f727dd71f1e8232f6171882`.
Native Save As regenerated root previews; cached Data9074 stayed exact at
`575645e2455199d7cc0c65fab8002b9e025765ba19b8b03c6e51c000f4915e89`.
Independent Apple toggles confirmed the native delta is Node field 18 plus one
field-7 and one field-42 membership entry, while field 20, cache data, and
global Show field 6 remain exact.

The post-cut gate passes 8/8 focused slide-number codec, 98/98 Keynote library,
7/7 focused visibility, 22/22 slide-preview, 9/9 `--features keynote` facade,
and 7/7 doctests. Keynote all-target checking, strict Keynote library
Clippy/rustdoc, host library check/no-run and examples, formatting, and diff
checks are green. The fuzz target compiles and completes a bounded 16-run
stable control-flow smoke; missing sanitizer symbols mean it is not
sanitizer-backed fuzz evidence. The boundary unit suite passes 138/138; live
slide-number host, placeholder host, and focused audits are clean. The full
checker reports only the unchanged 14 dependency-policy baselines. Native and
exact-artifact compatibility gates are final.

## 2026-08-11 amendment: focused Keynote soundtrack-settings exit

Soundtrack playback settings have crossed the format boundary into canonical
`soundtrack::{Mode, Settings, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` plus the Package read/edit/apply methods. This supersedes the older
claim that all soundtrack mutation remains in the monolithic editor, but only
for optional volume and mode. Soundtrack item/media CRUD, creation, allocation,
and reclamation remain distinct host responsibilities.

The focused owner proves the unique Document field-2 -> Show field-17 ->
type-21 Soundtrack rooted chain and its exact aggregate/field reference
metadata. It rejects zero/external/aliased references, duplicate selected
messages, merge/diff state, malformed component framing, and changed
non-exact/legacy provenance. A missing soundtrack is readable as absence but
cannot be synthesized by this edit.

Strict raw decoding owns canonical field-1 fixed64 volume, field-2 signed mode,
and streamed nonzero field-3 data references. It forces and cross-checks a
scalar Buffa lazy projection, then proves field-3 order against aggregate and
field metadata plus PackageMetadata component/data ownership and unique safe
`Data/` members. The generated closure is five files/27,753 bytes, zero
repeated views, within 32 KiB, with digest
`458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3`.
There is no generated encoder or repeated media collection in production.

The codec's byte/field/work/nesting/media report is merged with bounded rooted
graph, references, metadata, component, compression/output, reassembly, reopen,
and exact-delta work. Fallible allocation and content-redacted typed errors are
part of the publication contract, but this amendment does not declare the
shared memory/work/output review complete.

No-op publication shares exact source and skips rewrite/reopen. A changed
commit rewrites one soundtrack component, reopens the candidate, and verifies
that only field 1, field 2, the selected message length, and corresponding
selected ZIP CRC/size/offset bookkeeping differ. Apply authorizes exact source
and retained target; inverse restores exact bytes.
Patches remain process-local and output is through `write_to`.

This playback-only cut preserves all root previews, slides, ViewState,
slide/node caches, soundtrack field-3 ordering and bytes, PackageMetadata and
data-reference declarations, data assets, and unknown records. The retained
IWA item APIs, `KeynoteSoundtrackItemInfo`, `soundtrack_items` module and
example/tests, creation paths, and their necessary shared wire/media code are
not retired.

Compatibility was exercised from Apple-resaved 506,640-byte
`69795554212651b261f5ffd71dd5cf511544f285cab680d724a9de7d3f04b14d`.
Rust's same-size Loop/0.35 candidate
`6367e38a2edeebe6e65b148d0fd2aae555ee219dc1a65c339954047eb533ce1a`
changed only `Index/Document.iwa`; inverse restored the source. Keynote opened
warning-free, displayed Loop/0.3499999940395355 and `ringin` 00:00:01, and
played it. Native Save As produced 506,651-byte
`e264f4e714b0c44fca420b2c7b43e18f2ed1be99a766d25fe901f68d5f8bc299`.
The media file remained exact at
`5a08f48c4f86074e14a763d4f19f49ca31196a7a5f52fb48960e76b6f3d3d96b`,
the slide and three previews remained exact, and the normalized post-native
restage was a byte-exact no-op.

The exact exit deletes
`KeynoteEditor::{soundtrack_settings, set_soundtrack_settings}`, the whole
68-line `soundtrack.rs` editor source/module, settings-only
`patch_soundtrack_wire`, the dead decoded-native soundtrack-record field, two
whole settings tests and exclusive support (157 test lines), and the 29-line
direct mutation example. Production changes by +2/-91 lines. The inspector and
README migrate to the focused Package. The soundtrack-item CRUD/module,
`KeynoteSoundtrackItemInfo`, shared wire/media paths, creation, lifecycle,
example, and tests remain; this is not an item/media host exit.

No manifest dependency changes. Debt 014 (`litchi-iwa -> litchi-keynote`)
remains, as do 64 workspace packages, 235 internal declarations, 14
`litchi-iwa` dependency declarations, and 14 ordered debts.

The exit gate passes 5/5 soundtrack codec, 1/1 focused scaling unit, 4/4
focused settings, 99/99 Keynote library, 10/10 `--features keynote` facade,
and 8/8 doctests. Keynote
all-target, warning-denied Clippy/rustdoc, focused and retained examples, host,
formatting, and diff checks are green. Performance review found no P0/P1 issue;
the test-only `media.rs` regression drives realistic 4,096/8,192
metadata/media states through the real streaming path. Reference count doubles
exactly and fields/work/references stay within 2.3x; no wall-clock performance
claim is made. Boundary tests pass 152/152, host and focused audits each report
zero diagnostics, and the full checker reports only the unchanged 14 baselines: six
dev-only annotation findings and eight edge classifications.

## 2026-08-11 amendment: Numbers sheet-order host exit

The focused exit for sheet ordering is
`sheet::order::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` with
`Package::{edit_sheet_order, apply_sheet_order}`. One semantic selector and
one checked final-position destination replace direct editor movement; existing
semantic sheet iteration remains the read surface. The transaction exposes no
native identifier, component, protobuf, or source artifact.

This vertical owns both order lists, not merely Document field 1. Document
field 5 selects a same-component type-205 sidebar root whose field-2 children
must correspond positionally through their field-3 sheet associations. The
Document field-1 sheet references and sidebar-root field-2 child references
must be unique ordered aggregate subsequences and move together. Any selected
order reference attributed through a `FieldInfo` is refused; sidebar, child
association, and descendant metadata is accepted only on exact field 5/3/2
paths. Roles must be nonzero, nonexternal, disjoint, uniquely resolved,
canonical, non-merge, and within `Index/Document.iwa`. Only plain type-2
`TN.SheetArchive` changed sources are proven; FormBasedSheet and split ownership
return `UnsupportedSource`.

The intentionally scalar-only
`TNNumbersSheetReferenceArchive.proto` projects `TSP.Reference`. Handwritten
strict two-pass routing owns Document 1/5 and TreeNode 2/3 repeated records and
forces Buffa parity per scalar. No generated repeated view or encoder owns
preservation. The build produces five files/32,579 bytes, zero repeated-view
types, below 33 KiB, SHA-256
`2a0850fd82cfbf337ed48e582d4a998bd27e5046eb63c61f6939fa5ff1a09854`.

Finite codec and transaction budgets cover bytes, fields, depth, wire work,
references, object-index lookup, object/message/field metadata, raw and
aggregate reorder, archive extent/allocation, compression/output, reassembly,
preview deletion, reopen, and exact physical comparison. Fallible allocation
and content-redacted typed failure are atomic.

Same-position no-op shares exact source and skips native resolution/reopen. A
changed source must contain exactly one of each canonical root preview;
missing/repeated preview states are unsupported. Commit rewrites one component,
deletes all three previews, reopens once, and proves the exact dual-order delta.
Forward apply verifies previews 3 -> 0 and inverse verifies 0 -> 3. Changed
apply authorizes exact source and stored target, charges both plus reopen work,
and verifies the moved identity; conflict and inverse remain exact. Changed
legacy/non-exact sources fail closed. Patch remains process-local and output
uses `write_to`.

The focused cut preserves child IDs/nodes/associations/descendants, CalcEngine,
ViewState, sheet/table/drawable graphs, global table order, data sidecars, and
unknowns. It does not retire sheet creation, duplication, removal,
FormBasedSheet/general Document-reference helpers, table/drawable CRUD,
component/ID allocation or reclamation, or their mixed tests/examples. The
host deletion must therefore remove only direct ordering scope and retain that
shared substrate.

The independent P0/P1 review found no release blocker and no O(S²). A 4,096 to
8,192 reference test passes the strict codec, raw record reorder, and core
aggregate-header reorder and holds production work/references/payload within
2.3x plus a fixed 32-unit allowance; codec-only scaling is strict 2.3x. No
timing claim is made. P2 tradeoffs remain the bounded roughly-four-snapshots-per-sheet
Patch for no reselection/O(1) inverse, transient Vec-to-Arc target duplication,
and one possible bounded O(package-bytes) byte-equal-source authorization
comparison before charging; identity authorization is O(1).

Matched Apple artifacts are control
`f9c5cbec4f422484c63d1d39bd8d09da122d011596561a5feb2ad1e812574990`
(133,594 bytes) and reorder
`7b3bcbc853346a433e84ee815d28671d01fc3da857e43b8b7d29b310f94e7e1a`
(153,498 bytes). Native reverses both order lists/aggregate subsequences and
keeps child field-3 associations exact; 93/103 decompressed members, including
all table sidecars, are unchanged. Apple TableInfo cache culling, physical
subgraph movement and tree/ViewState/ID/metadata/property/timestamp churn are
normalization rather than focused requirements.

Rust artifact
`97c76894503a2628c1828babd93d9a9a891794d86c86177cab60f09333997a68`
opened in Numbers 14.4 without warnings, repair, or conversion and retained the
expected `FirstCreated`/`SecondCreated`, `A-new`/`A-old`/`B-only` associations
with benign CalcEngine preservation. Native Save As/close/exact reopen yielded
the same semantics at 103-member
`4aa257e4db61a3c03950360b29267c9495985d460ae22b6f679bee31f2693217`
and regenerated all three previews to the matched Apple hashes. Focused
same-position no-op and inverse were byte-exact at that hash with 0/0/0/false
diagnostics.

The format implementation is five sources: public `sheet/order.rs`, transaction
`package/sheet_order.rs`, and the frozen private
`package/sheet_order/{error,resolve,rewrite}.rs` tuple. The exact host deletion
removes `NumbersEditor::move_sheet` plus exclusive `sheet_index` at -58
production lines, changes tests +2/-43, and deletes the 23-line legacy move
example. The retained remove-sheet example migrates +2/-6 to a semantic
selector. Sheet add/duplicate/remove and shared substrate remain.

Codec/protobuf tests pass 7/7 and 132/132; Numbers passes 109/109 library, 4/4
private sheet-order, and 1/1 public integration tests. Boundary tests pass
165/165, Python compilation/diff are green, and live host/focused audits each
report zero diagnostics. The full checker has only the unchanged 14 baselines:
six missing dev-only `soapberry-zip` annotations and eight unclassified edges
(the same six plus `litchi-odf-common -> xml-minifier` and
`litchi-opc -> xml-minifier`). No dependency edge closes. Debt 014
(`litchi-iwa -> litchi-keynote`) remains, with topology at 64 workspace
packages, 235 internal declarations, 14 `litchi-iwa` declarations, and 14
ordered debts.

## 2026-08-11 amendment: Numbers table-title host exit

The direct Numbers table-title seam moves to
`table::title::{Settings, Edit, Patch, Commit, Diagnostics, Error, LimitKind,
Path}` and `Package::{table_title_settings, edit_table_title,
apply_table_title}`. Semantic sheet/table selectors replace host raw IDs.
The focused signatures expose no raw source, component, or generated value;
source bytes remain crate-private and output is through `write_to`. Optional
field-22 visibility and field-37 outline presence are lossless, and
consuming `Edit::set` stages a complete value without touching the package.

Changed ownership follows the focused rooted Document -> Sheet/FormBasedSheet
-> TableInfo -> TableModel chain and exact reference metadata, refuses locked
tables, and requires canonical nonaliased rendering prerequisites before
publishing a visible title: finite nonnegative field-33 height, field-30
paragraph style/type 2022, and field-36 shape style/type 2025. Unsupported
style or ownership prerequisites fail closed. Changed admission scans
`Index/ViewState.iwa` and rejects any native type-6284 table-name-selection
message with `UnsupportedSource`; the transient selection state is an
unsupported dependency, not a write right. Reads and exact no-ops retain broad
compatibility. Every other ViewState byte in an accepted changed source remains
outside the write set and exact.

Strict raw routing validates the three title scalars and the two reused scalar
references before forcing/cross-checking Buffa views. There is no generated
encoder or repeated view. The deterministic five-file closure is 32,332 bytes
under 33 KiB with SHA-256
`56cfd70666ffa6079175bdab0a63a4ddd055099edf3c771ed3ad8b3051596ee1`;
codec/protobuf verification is 9/9 and 141/141.

No-op shares exact source and skips native rewrite/reopen. A change rewrites
only the selected TableModel in `Index/CalculationEngine.iwa`, removes each
existing canonical root preview (zero to three), and reopens/verifies the exact
delta. Apply authorizes exact source and target artifacts, conflicts on drift,
and inverse restores exact source bytes and previews. Accepted ViewState and
every nonselected component remain exact; changed legacy/non-exact provenance
is unsupported.

The Numbers 14.4 control resave is 136,204 bytes,
`25c9fc858ca4fb4f1fedeafb944e96afb81af03a082a41be297ecf6f2542dbdb`;
the title-hidden resave is 136,273 bytes,
`ac8a7117ad6256b0da2e6d191b9e64f721b689d71696a89ac0f78bc6aa513a28`.
Native hiding removes raw field 22 rather than writing false, while field 37
remains independently presence-sensitive. This oracle does not authorize a
type-6284 ViewState rewrite; changed admission rejects that transient state.

Rust starts from the 136,357-byte exact source
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
produces the 136,351-byte hidden candidate
`4c7f6340b6f2675240577c5b59d5c154de24c8a7e763a31257c56a9899a8e40c`,
and inverses byte-exactly to the source. Numbers 14.4 opened the candidate
without warning, showed Table Title unchecked, retained its 22-by-7 shape, B2
fixture marker, and B3 value 42, and preserved those semantics through Save As,
close, and warning-free exact-URL reopen. The 136,353-byte native resave is
`5b162f8431f45333f0ae9a8654dfa724794f2ec2b391ea11f6a5eee7822cbb10`.

Final performance review is P0/P1-clean. The real rooted Package path from
4,096 to 8,192 objects records fields 53,307 -> 108,363 (2.0326x), `WireWork`
315,936 -> 636,752 (2.0155x), references 16,386 -> 32,770 (exactly `2 + 4N`,
1.9999x), and `TransactionWork` 9,084,384 -> 18,298,157 (2.0142x), all at or
below 2.3x. A maximum-minus-one work budget rejects before output. The only P2
costs are linear selector temporary vectors and redundant decode passes for a
changed edit; no timing claim is made.

The exact host cut removes two public NumbersEditor methods, 32 production
lines, 245 direct-test lines, and `edit_numbers_table_title`. The private
Numbers package helpers and wire path remain because Pages and Keynote still
use them; their table-title APIs and table CRUD are not claimed by this exit.
Boundary regressions pass 173/173. Final gates pass 111/111 Numbers library,
2/2 private table-title, 5/5 public table-title, 9/9 codec, and 141/141 full
protobuf tests. The full checker retains only 14 unchanged dependency-policy
baselines. No manifest edge or ordered debt closes: the current inventory is
64 workspace packages, 237 internal declarations, 14 `litchi-iwa`
declarations, and 14 ordered debts, including debt 015.

## 2026-08-11 amendment: aggregate Pages section-settings host exit

Deletion gate 3 now assigns the complete existing-section settings vertical to
`litchi-pages`. The archive-free `section::Settings` and focused
`section::settings::{Edit, Commit, Patch, Diagnostics, Error, LimitKind, Path,
DependencyKind}` surface replace raw-ID host access for native fields 17--22,
26, and 28. `Package::{section_settings, edit_section_settings,
apply_section_settings}` resolves exact names or checked semantic positions;
it exposes no monolith, archive, protobuf, Buffa, component, object-identifier,
wire, or exact-artifact type.

This is one physical-writer transfer. The retained `section_name` and
`section_pagination` APIs are projection-scoped ergonomic facades over the same
aggregate transaction core, not compatibility implementations in the
migration host and not independent package writers. Their previous semantic
and native evidence remains valid, while earlier claims that fields 20--22 and
field 26 have separate current physical mutation owners are superseded.

Strict raw preflight and a private aggregate Buffa lazy view agree on all eight
selected optional fields. The borrowed view retains no unknown/repeated state
and cannot encode; caller-owned records retain preservation authority. Fields
23--25 are target-sensitive template prerequisites, while fields 29, 30, and
31 and every unknown record remain outside the settings mutation. Missing
otherwise-valid prerequisites fail as typed unsupported dependencies;
malformed, ambiguous, aliased, or contradictory ownership fails closed. Final
generated code is exactly five files and 80,202 bytes under 80 KiB, contains no
repeated lazy view, and has aggregate SHA-256
`2202f4b1d394346450cb9f88a41c2784ab476cff23b181fffbab6f37b4a42b62`.
The focused protobuf suite passes 149/149.

Exact no-ops share the source and avoid dependency/cache scans, preview
planning, reassembly, and reopen. Changed publication rewrites one selected
section payload while preserving the unique rooted layout/cache edge, metadata,
and every root preview exactly, then reopens the whole candidate under retained
limits. Template/background payloads, opaque and detached cache objects,
sibling sections, and unrelated components/members remain exact. Patch
application is authorized by exact artifacts, conflicts fail, and inverse
application restores the accepted source byte-for-byte.

The host cut removes `PagesEditor::section_settings`,
`set_section_settings`, and `set_section_name`; the raw-ID
`set_pages_section_settings` example; duplicate host tests; and stale README
usage. Background mutation remains a separate host capability with only its
private payload helpers retained or relocated. Changed legacy nested packages
are no longer normalized by the deleted settings/name path; reads and exact
no-ops remain broad, and changed focused transactions return
`UnsupportedSource`.

Matched native Pages 14.4 evidence now proves fields 17, 19, and 28 independently
as one exact false-to-true scalar delta on object 1732889/type 10011. It keeps
field 18, message header/references, templates, header/footer storages, entry
names, caches, and previews exact and reopens warning-free with the expected UI
behavior. ADR 0008 records all seed/control/change hashes. The independent
production gate records 77-to-77
selected fields, 564-to-564 strict wire work, 4-to-4 references, and
292,154-to-587,222 `TransactionWork` (2.0100x) when rooted real-package objects
double from 4,096 to 8,192. Output allocation and reopen counts remain one at
both sizes; a maximum-minus-one work ceiling refuses before output with both
counts zero. Focused integration passes 7/7, four private production/security
tests cover budget observation, object scaling, alias-metadata refusal, and
repeated-reference scaling/max-minus-one refusal; the projection suite passes
149/149, and locality review is clean. The full Pages library/integration gate
is 118/118; boundary regressions pass 181/181; focused facade/host audits report
zero; and the live
checker retains only 14 unchanged baselines. The matched native pairs are the
UI oracle; no separate Rust-authored application artifact is claimed.
The transfer closes no manifest edge and deletes no ordered debt: topology
remains 64 workspace packages, 237 internal declarations, 14 `litchi-iwa`
declarations, and 14 ordered debts, including debt 017. Remaining Pages
editors, examples/tests/fuzz ownership, durable patch serialization, atomic
filesystem publication, shared aggregate memory/work completion, remaining
Buffa conversion, and the other host debts continue to block monolith deletion.

## 2026-08-11 amendment: Numbers table-cell read owner

The monolith exit now includes one narrowly completed read seam.
`litchi-numbers` owns `table::cells::{State, Storage, Error, LimitKind, Path}`
and selector-first `Package::{table_cell, table_cells}`. One coordinate is
checked directly; a half-open range returns a bounded, fallibly allocated,
dense row-major result. `Storage::Missing` and
`Storage::Stored(Value::Empty)` are separate semantic states, and name
ambiguity, bounds, element limits, text limits, and allocation failures are
typed without exposing authored content or native identifiers.

This is not yet the physical table-cell exit. The methods consume the
already-eager semantic `Table` built by `litchi-numbers::package::extractor`
through the existing BNC/protobuf decoder. The new strict storage/dependency
Buffa codecs are preparatory and do not power this read path. Earlier text that
assigned all BNC-backed semantic cell reads to `litchi-iwa` is superseded, but
the host retains its cell mutators and helpers, formula compiler and AST wire
handling, calculation-engine mutation, downstream cache updates, and
publication. Consequently no legacy editor method, test, example, module, or
source file is retired here.

The range algorithm is bounded analytically: for area `A`, selected-row-span
materialized cells `K`, selected owned-string bytes `B`, selected strings `T`,
and table materialized-cell count `M`, non-empty work is
`A + 2K + 2*O(log M)`, with one `A`-state vector and `T` fallible string
allocations totaling `B`. Empty ranges perform no scan or allocation.
4,096-to-8,192 paired cases keep size-sensitive terms at or below 2.0x and
result allocations at one; element/text over-limit cases fail before the
result allocation. This is not latency or RSS evidence.

The 136,357-byte native `basic.numbers` oracle, SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
establishes Sheet 1/Table 1 dimensions 22x7, stored text at B2, stored number
42 at B3, missing A1, dense row-major A1:C3, and out-of-bounds A23. The
140,498-byte formula/rich-text oracle, SHA-256
`80deb7b87df27f58b26e6f247acee9d1fc6dcd3d268e85046c3efc16070b2edf`,
adds formula and rich-text-backed reads. Neither proves native stored-empty
presence; that distinction is synthetic-test evidence. No native write,
cache, preview, reassembly, or Save claim follows.

Gates pass 114/114 Numbers library, 4/4 public read integration, 13/13 strict
codec, 163/163 full protobuf, and 187/187 boundary tests, plus strict
library/test Clippy and warning-denied rustdoc. The full checker retains only
14 unchanged baselines. This slice removes no host declaration, manifest edge,
or debt: topology stays at 64 packages, 237 internal declarations, 14
`litchi-iwa` dependency declarations, and 14 ordered debts.

## 2026-08-12 amendment: Numbers table-cell mutation exit slice

The physical exit has begun without invalidating the read-only history above.
The focused Numbers owner now presents `table::cells::{Input, Change, Edit,
Patch, Commit, Diagnostics, Error, LimitKind, Path, DependencyKind}` through
selector-first `Package::{edit_table_cells, apply_table_cells}`. It owns one
bounded final-overlay transaction for supported scalar/string/tile/rich/cache
changes and exact directional publication. It does not expose or relocate the
legacy raw-ID editor facade.

The accepted slice is scalar set/clear, direct/unsegmented string-list
ownership with exact refcounts, missing sparse-tile growth including a
synthetic 513-row boundary for finite non-text scalars, in-place authored-text
replacement in uniquely owned rich backing while retaining key/storage
identity and releasing exact style references, and downstream cache refresh
for a strict supported formula graph evaluated against the complete batch. It
does not migrate formula compilation/AST construction, formula/error-cell mutation,
general rich text, formats, controls, comments, table topology, or the
Pages/Keynote attached-table paths. CalculationEngine field 14 is projected
and its rooted HeaderNameMgr reference validated; only the referenced manager
payload/update semantics are unprojected, so a manager-backed header change
refuses as `HeaderNameIndex`. Sparse text to a missing tile refuses as
`SharedString`. Segmented string lists, existing formula/error cells, shared
rich text requiring a FieldInfo reference transition,
noncanonical/ambiguous FieldInfo rich ownership, modeled missing storage
prerequisites, and unsupported formula dependencies fail as
`UnsupportedDependency` (missing
storage as `CellStorage`). Impacted active merge, pivot, category, spill,
hidden, or conditional-style state refuses by its matching kind while
unrelated/inert state remains exact. Malformed routes fail as `InvalidSource`;
an unmodeled stored BNC value/source kind fails as `UnsupportedSource`. Locks
likewise fail typed and before publication.

Canonical payload field-1-to-storage and storage field-2-to-style FieldInfo
metadata may be present on the unique rich path and remains exact when no
field-specific reference transition is required.

The private strict storage and dependency projections are respectively five
files/465,932 bytes/SHA-256
`1a894fd5d22b004db664bc7c348d9591a4608ab9263a8122c726c8a1ecb0c3b3`
and five files/544,538 bytes/SHA-256
`2fba7c22aef58ed3cfe6eba1f77e5eaf79d2597dd79966e05d20e50c0e2b33b3`,
with no generated repeated views. The dependency-only formula projection is
five files/201,539 bytes/SHA-256
`ccd972b3dcd76b6142342d36435f2f76a305c029265853ced04d64c1e2bf1752`;
its focused codec passes 7/7 and the full protobuf suite passes 178/178. Raw
records remain authoritative. The PackageMetadata projection is five
files/145,681 bytes, generates no repeated view, and has SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`;
its focused gate passes 7/7. Changed commits verify exact message/reference
deltas and preview membership, then reassemble/reopen once. Exact apply borrows
the patch and reuses a private retained source/target `PackagePair`; this
avoids a second reopen but leaves
process-local memory, serialization, versioning, composition, merge, and
history debt.

Reads and exact no-ops remain broad. A changed package without an exact
physical `SourceCatalog`, including nested legacy layout, fails as
`UnsupportedSource`.

Final rooted 4,096-to-8,192 transaction-work ratios are 1.1899x numeric,
1.2245x unique text, 1.1396x same-tile, and 1.8021x formula, with governed
subterms no worse than 2.0x. Required-minus-one formula/sparse cases stop with
zero component, reassembly, output, reopen, and locality work. Numbers 14.4
passes the exact numeric B3=43 scalar and unique-rich C2 no-impact
commit/apply/inverse and
open/Save As/reopen gates. The latter preserves its independent formula/cache;
it is not impacted-formula-refresh proof, and unsupported impacted formulas
still fail as `FormulaCache`.

The completed exit removes `NumbersEditor::{set_cell, set_cells, clear_cell}`,
the Numbers-only raw-ID model set/set-batch helpers,
`TableCellBatch::apply_numbers`, 15 obsolete direct tests, and the legacy
example. Shared `TableCellBatch::{collect, is_empty, len, apply_attached}`,
attached/package helpers, storage/wire/cache/formula machinery, Pages/Keynote
paths and builders, and fixture-only test adapters remain. Numbers gates pass
237 library tests with 4 ignored and 15/15 public cells; boundary regressions
pass 196/196 and the host/focused source audits are empty. The focused
implementation adds `litchi-numbers -> litchi-iwa-text-wire`; topology is 64
packages, 238 internal declarations, 14 `litchi-iwa` declarations, and 14
ordered debts. The Numbers host debt remains because the retained editor is
much broader than this cell slice.

## 2026-08-12 amendment: Keynote existing-slide deletion host exit

Deletion gate 3 advances for one structural Keynote operation. The concrete
package now owns deletion of one existing, non-final slide through
`slide::delete::{Edit, Patch, Commit, Diagnostics, Error, LimitKind, Path}` and
`Package::{edit_slide_deletion, apply_slide_deletion}`. The selector is an
exact navigator name or checked position; the host's raw-index method is not
retained as an alias or shim.

The focused owner admits only a unique flat rooted
Document -> Show/SlideTree -> SlideNode -> Slide path. Exact aggregate
references and any optional field attribution, package-wide exclusive inbound
ownership, one selected Node message, one selected Slide message, absence of
merge/base/diff state, and exact current PackageMetadata
component/UUID/external-edge/data-owner facts are all changed-operation
preconditions. Unsupported hierarchy,
deprecated or secondary slide roots, duplicate or cross-component ownership,
versioned records, aliases, contradictory locators/counts, and malformed
metadata refuse before publication.

The accepted transition removes the selected Show reference, Node and Slide
objects, their two object-to-UUID records, the one unversioned Node-component
external reference to the Slide object when that ownership form is present,
and the exact data-owner/count entries attributed to those two objects. The
supported component-level edge remains instead. Component registrations, the
last-object identifier, co-located objects, global data-catalog records, data
payloads, and all unrelated PackageMetadata fields remain. A component
data-reference record remains with surviving owners or is removed with its
final owner.
Exact root previews are invalidated, one package is reassembled and reopened,
semantic order is read back, and the exact inverse restores the source.

This exit deliberately supersedes the old host's orphaned-media-reclamation
behavior. Slide deletion is not package GC: no slide component, global
data-catalog record, `Data/` member, or media payload is reclaimed. Shared,
uncertain, or newly unreachable media remains preserved. A later GC owner must
establish a separate complete reachability/disposition proof and native
evidence before it can remove any such content.

The host cut removes `KeynoteEditor::remove_slide`, the entire
`keynote/editor/slide_delete.rs` source and module declaration, the direct
`remove_keynote_slide` example, and obsolete direct deletion assertions. The
retained source-free regression is creation-only and does not claim its
child-to-parent-slide backlink topology is deletable; the focused transaction
correctly refuses that surviving owner as `AmbiguousOwnership`. No
compatibility method or public bridge alias replaces the retired host method.
This retires direct existing-slide deletion only; creation, insert/duplicate,
layout, drawable, chart/table, media, soundtrack-item, and remaining editor
ownership stays in the host.

Accordingly debt 014 (`litchi-iwa -> litchi-keynote`) and its manifest edge
remain. The boundary suite passes 204/204; focused and retired-surface audits
each report zero findings, and the full checker reports exactly the 14
established unrelated findings. Native hashes/UI observations and
PackageMetadata generated evidence are frozen in ADR 0008. The final topology
is 64 workspace packages, 238 internal dependency declarations, 14
`litchi-iwa` dependency declarations, and 14 explicit ordered debts. Keynote
passes 235/235 all-features tests and 9/9 doctests; the retained host library
passes 1,418/1,418, including permanent atomic refusal coverage for the
generated child-to-parent-slide backlink. These results close the focused
existing-slide deletion gate. The broader monolith exit and debt 014 remain
open, and the 14 established unrelated full-checker findings are unchanged.

## 2026-08-12 amendment: Numbers formula-cache prerequisite, not exit

The bounded internal cache planner preserves unrelated cycle markers
byte-for-byte, refuses when an impacted marked formula survives the final
same-batch overlay, and succeeds when that overlay removes it. Exact graph-work
max-minus-one refusal coverage, together with bounded scratch and allocation,
reduces prerequisites without moving the production surface.

The formula exit gate is not satisfied: public focused authoring has not
landed, production host formula setters and raw formula vocabulary remain, and
no dependency edge or ordered debt closes. This amendment accepts no native
formula-authoring or formula-authoring performance evidence.

## 2026-08-13 amendment: Pages section-background host exit

Deletion gate 3 advances for direct field-30 backgrounds on existing Pages
sections. `litchi-pages` now owns selector-first read, set-solid, clear,
exact-source patch application, and inverse through its focused
`section::background` transaction. The host's `PagesEditor::section_background`
and `set_section_background`, their two private implementation modules, direct
raw-ID example, duplicate host regression, and stale README usage are removed;
no compatibility alias bridges back to the migration host.

The transaction classifies absent direct fill, supported sRGB/Display-P3 solid
fill, and a byte-free `Unsupported` preservation state. It does not author,
clear, or normalize unsupported gradient/image/future fills. Duplicate,
wrong-wire, malformed, ambiguous, or reference-owned field-30 state fails
closed for changed operations. Changed publication preserves the selected
section's non-background fields, sibling sections, unrelated objects and ZIP
members, and the rooted cache/layout/preview state; it reopens before
publication and exact inverse restores the accepted source.

The private Buffa lazy projection is bounded and read-only: five generated
files, 99,593 bytes, zero repeated views, with the frozen aggregate SHA-256 in
ADR 0008. Codec and focused integration gates pass 8/8 each; the deterministic
4,096-to-8,192 work gate stays within 2.20x, fixes output allocation/reopen at
one each, and proves an instrumented transaction-work max-minus-one refusal.

This transfer does not close the Pages manifest edge or its ordered host debt.
Section text/templates/header-footer content, section lifecycle, tables,
drawables/media, broader editor paths, durable patches, and atomic library save
remain open. Native Pages 14.4.1 acceptance is now established for the two
supported field-30 transitions only: it opened both Rust candidates without a
repair/conversion prompt, showed page 1 dark-red `Color Fill` after replacement
and `No Fill` after clear, and retained those states through Save, close, and
exact-path reopen. The original candidates and exact inverses, as well as the
Pages-resaved ZIP hashes and post-resave no-op proof, are frozen in ADR 0008.
Pages' own resave is not asserted to retain byte or member locality.

## 2026-08-13 amendment: Keynote reader host exit

The Keynote read-facade cell exits the monolith. The complete 933-line
`keynote/document.rs` implementation, its module declaration, and its
`KeynoteDocument`/`KeynoteDocumentStats` exports are removed. That retires the
second package capture, object index, semantic cache, eager-Prost show/slide
decoder, graph resolver, and wide text extractor. The focused package still
uses six bounded, preflighted Prost decodes during semantic traversal; the
exit claim is removal of the duplicate eager host reader, not a Prost-free
focused implementation.

`litchi_keynote::Document::{open, open_with_options}` owns the surviving
semantic path API for complete ZIPs and frozen app-authored package directories.
It captures checked `PreparedSource` components and eagerly publishes an
archive-free full show, rooted text, source-derived metadata, and source
statistics. For a source-backed `Document`, metadata and statistics are
present; metadata combines semantic Show values with narrowly decoded
canonical-properties scalars when that diagnostic exists, so `Some` does not
prove sidecar presence. `litchi_keynote::Package` remains the exact regular-file
artifact owner for path/byte ingress, semantic projection, cheap shared
`semantic_snapshot`, exact `write_to`, and edit provenance. The package-derived
semantic snapshot is intentionally diagnostic-free, with `metadata()` and
`stats()` both `None`. Direct
`Package::open` refuses directory backing because an `Index.zip` or loose
`Index/` is not the complete artifact required for those operations. The
cross-format coordinator can delegate semantic reads through the same focused
boundary. The removed `from_archive_bytes` name was a direct alias for
`from_bytes`; the removed application statistic was always the constant
`Keynote`. No compatibility alias preserves either redundancy, and semantic
directory snapshots do not promise preservation of other sidecars, `Data/`,
previews, or exact package bytes.

This is capability parity with deliberate semantic correction, not structural
equality. Focused text follows rooted presentation reachability, slides retain
rich storage fragments instead of flattened legacy body/date text, and focused
metadata/validation are richer and stricter. Canonical logical
`Metadata/Properties.plist` lookup ignores hostile basename-only near-names.
That diagnostic has a centralized 64 KiB hard admission ceiling independent of
broader entry limits, and only public-metadata scalar fields are decoded.
The generated roundtrip, host compile/lint/doctest, focused path, native
fixture, and boundary gates cover the retired surface and replacement paths.
The fixed Apple-authored read-only fixture is
500,058 bytes with SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
The live retired-reader audit reports zero findings, and the full checker
continues to distinguish its dependency-policy baseline findings. Permanent
path tests prove packaged/directory semantic parity through both focused
`Document` and the coordinator, and match the directory snapshot to the focused
ZIP reader.
Frozen ingress and semantic gates pass archive-directory 16/16, detection
18/18, focused Keynote native 7/7, coordinator `iwork_path` 7/7, and metadata
scalar/64 KiB-cap unit coverage 1/1.
Keynote 14.4 read-only acceptance on an isolated copy showed the one expected
slide and its title/body/date sentinels without repair, recovery, or conversion
warning. Separately, the non-UI focused native-fixture gate reports one slide
and 959 objects. Native autosave normalized only the
disposable copy, so no claim is made that native open is byte-inert.

The monolith exit remains incomplete. `KeynoteEditor` and
`KeynoteDocumentBuilder` still live in `litchi-iwa`, so the manifest edge and
ordered debt 014 remain unchanged.

## 2026-08-13 amendment: Pages reader host exit

The Pages read-facade cell exits the monolith. The complete 478-line
`pages/document.rs`, `pages::document` module, `PagesDocument` re-export, and
`PagesDocument`/`PagesDocumentStats` types are removed together with their
duplicate Bundle, object index, root/body decoder, and package snapshot. No
compatibility alias preserves the retired reader.

Its capabilities move to the focused crate with an explicit provenance split.
`litchi_pages::Document` owns semantic ZIP and checked-directory path reads on
supported path-ingress platforms, borrowed/shared ZIP byte reads on every
supported platform, eager archive-free snapshots, text and sections,
source-derived metadata/statistics, and semantic validation.
`litchi_pages::Package` owns exact regular-file and byte ingress, source bytes,
package metadata/statistics, physical validation, edits, and the
existing archive-byte alias. It refuses directories. A package-derived or
semantically constructed `Document` has no source diagnostics, and a
directory semantic read makes no preservation promise for unselected
metadata, `Data/`, previews, media, unknown sidecars, or complete bytes.

This supersedes the earlier directory statement that all `Metadata/` was
outside the frozen adapter. Pages semantic ingress now captures exactly
`Metadata/Properties.plist`, `Metadata/BuildVersionHistory.plist`, and
`Metadata/DocumentIdentifier`, retaining at most 64 KiB from each, from the same
captured authority as the components. It still excludes every other
sidecar and cannot enter preserve-mode editing. Packaged and directory
semantic reads therefore share metadata semantics without conflating either
with exact-artifact ownership.

Deletion gate 3 advances for the retired read behavior with deliberate
correction: the focused owner uses native empty-root, section-table,
exact-name, rich-run, and UTF-16-boundary semantics rather than the retired
synthetic one-section view. Rootless fallback reproduces the retired 14-type
object trigger and registry-message aggregation, including source-order
newlines and empty fragments, after strict raw/Buffa validation and before
aggregate text publication. The public duplicate eager-Prost reader is gone.
Focused Pages is nevertheless not Prost-free or wholly lazy:
one bounded rooted StorageArchive decode remains behind strict raw/Buffa
qualification, while fallback candidates receive full known-field raw storage
validation before the Buffa text projection.

The native read-only gate is frozen in ADR 0008. Its isolated copy opened
without repair or conversion and retained the expected one-section semantics,
but Pages silently normalized that disposable artifact, so native open is not
claimed byte-inert. The tracked source remained unopened and exact.

This reader exit does not close the remaining source-boundary gates. Packaged
semantic ingress can still materialize irrelevant supported ZIP entries under
the generic source limits before dropping them. The three selected canonical
metadata authorities are nevertheless declared-size and
unsupported-compression preflighted before any package entry payload is
materialized for path, borrowed-byte, and shared-byte ZIP ingress. They match
exact raw logical-name bytes after stripping only the selected legacy
outer-package prefix, exclude raw near-names, and require local/central ZIP
names and methods to agree. The focused
semantic opener exposes the content-free `ReadError` taxonomy rather than
lower-layer diagnostics. Windows Pages file and directory path ingress fails
closed until descriptor-relative, reparse-safe stable identity is available;
borrowed/shared byte ingress remains supported. The unrelated-member
materialization debt and deliberate Windows path capability gap prevent a
claim of complete Pages ingress or completion of the whole Pages portion of
deletion gate 3.

The final focused Pages gate passes 153/153, with supporting archive and
detector suites at 93/93 and 32/32 and the host generated-roundtrip at 1/1.
The boundary ratchet passes 227/227; the live retired-reader and focused
public-API audits both report zero findings.

The monolith exit remains incomplete. `PagesEditor`,
`PagesDocumentBuilder`, creation, and extensive chart/table/media/formatting
examples and tests still live in `litchi-iwa`; production host code still uses
focused Pages types. Ordered debt 017 and the `litchi-iwa -> litchi-pages`
manifest edge remain unchanged.

## 2026-08-13 amendment: Numbers reader host exit

The Numbers read-facade cell exits the monolith. The 460-line host
`numbers/document.rs`, its module and re-export, `NumbersDocument`, private
reader state and statistics type, and the 142-line reader-only `NumbersSheet`
adapter are deleted without aliases: 602 host reader lines in total. The cut
removes the duplicate bundle capture, object index, root/sheet/table walker,
repeated semantic conversion, validation path, and public
`bundle()`/`object_index()` escape hatches.

Supported read responsibilities move to focused owners by provenance.
`litchi_numbers::Document` owns archive-free semantic file/directory paths,
borrowed/shared bytes, and checked source and semantic options. Semantic zero
ceilings are exact and over-hard requests fail through `DocumentLimitsError`.
It owns rooted sheets and selectors, cheap snapshots, rooted plain text,
validation, and optional source metadata/statistics. `litchi_numbers::Package`
owns exact complete regular-file/byte artifacts, physical/storage diagnostics,
write/edit provenance, and focused transactions. A package-derived or
semantically constructed document has no source metadata or stats, and
`Package::open` continues to refuse directories.

Source-backed metadata captures only the three canonical Numbers authorities,
from the same frozen source as the semantic components and only after Numbers
classification. It excludes every other sidecar, `Data/`, previews, media,
unknown package members, and exact bytes. This advances frozen-directory and
semantic-reader coverage without turning an app-authored directory into a
preserve-mode package.
Unix path ingress uses pinned, no-follow capture; other non-Windows targets use
version-checked path capture. Windows path ingress remains an explicit
capability gap: both file and directory paths fail closed until stable,
reparse-safe handle traversal is available, while borrowed/shared byte ingress
remains portable.
Each selected authority is physically capped at 64 KiB before materialization;
a narrow `plist::stream` event projector applies fixed structure and
retained-scalar ceilings before source diagnostics are published, without
deserializing a general scalar DTO or plist value tree.

Deletion gate 3 advances with deliberate semantic correction. Rooted
`Document::plain_text` follows semantic workbook reachability and ordering and
works on the native fixture that the legacy reader rejected during document
construction, before public `text()` was reachable. Recovered private legacy
storage output matched `Package::text` on two frozen fixtures, but
`Package::text` remains a separate physical/storage diagnostic with no general
parity claim. Raw archive/object-index getters and
archive-aware sheet adapters are intentionally not replaced because they
violate the format-owned semantic boundary. The redundant archive-byte aliases
and constant `Application::Numbers` statistic are also removed rather than
carried forward.

The focused reader remains eager and substantially Prost-backed. Existing
private raw/Buffa projections cover selected ownership/scalar seams, while the
larger root, sheet, table, tile, list, comment, rich-text-sidecar, and formula
graph remains generated-Prost debt. Consequently this reader exit is neither
a whole-graph Buffa-lazy claim nor completion of the Numbers portion of
deletion gate 3.

Deletion gate 3 is frozen by 16/16 focused reader cases, with a seventeenth
Windows-configured case; 240 Numbers library cases pass and four are ignored;
compatibility and name gates pass 5/5 and 10/10. Archive coverage passes 127
cases (125 unit plus two integration), detector coverage passes 40/40, and the
host library passes 1,397/1,397, while generated-roundtrip and doctest gates
pass 1/1 and nine passed with three ignored. Host all-target check and no-run,
strict scoped host Clippy, focused
all-target Clippy, strict focused rustdoc, formatting, and diff checks pass. The
boundary units pass 237/237 and both live retirement/API audits report zero
findings. The host-scoped cut touches 15 files with 329 insertions and 888
deletions, net -559, including the 602 reader-owned source lines. Broad host
all-target Clippy remains blocked by unrelated existing lints; the global
boundary policy still reports 14 unrelated `soapberry-zip`/`xml-minifier` debt
findings.

ADR 0008 freezes native evidence from an isolated copy: Numbers 14.4 build
7043.0.93 accepted the one-sheet, one-table 22-by-7 workbook without repair or
conversion and agreed with focused semantics, but silently normalized the copy
on close. The tracked source remained unopened and exact. No byte-inert-open,
package-locality, or performance claim follows.

The monolith exit remains incomplete. `NumbersEditor`,
`NumbersDocumentBuilder`, `NumbersTable`, `TableDataExtractor`, creation,
editing, examples, and broad compatibility tests still live in `litchi-iwa`.
They continue to use `litchi-numbers`, so ordered debt 015 and the manifest
edge remain unchanged.

## 2026-08-13 amendment: Numbers dimension-size host exit

The Numbers-specific row-height/column-width facade leaves the monolith.
Focused `Package` owns selector-first read, edit,
source-bound patch apply, inverse, diagnostics, limits, strict reopen, and one
header-bucket locality. `Size::Default` retains the native absence/zero
override meaning; explicit positive finite `Points` remains a distinct state.
Raw IDs, archive names, generated messages, and physical bucket references do
not cross the focused public boundary.

The retirement changes five paths with four list-format insertions and 312
deletions, net -308. It removes the six public `NumbersEditor` size methods (65
source lines, 63 method-body lines), 200 test-section lines (about 197
test-body lines), the 41-line `edit_numbers_table_dimension` example, and the
host `Dimension`/`Points`/`Size` re-export names. Boundary ratchets keep those
names and facades retired without rejecting direct focused-package use.

This advances the Numbers portion of deletion gate 3 but does not complete it.
Pages and Keynote still depend on a private host `IWorkPackage` header-bucket
helper for their embedded tables, and the host still owns the shared table-
resize/axis-topology machinery. That helper remains explicit ordered debt; it
must not be recast as a public Numbers compatibility layer. Debt 015 and the
manifest edge therefore remain open.

The isolated native candidate opened in Numbers 14.4 build 7043.0.93 without
repair, recovery, or conversion UI. Row 5 remained 32 pt, column C 124 pt, and
preserved column F 98 pt before save and after save/close/reopen; the one-sheet,
one-table 22-by-7 topology, headers 1/1/0, `B2` marker, and `B3 = 42` remained.
The saved 136,615-byte 43-entry ZIP has SHA-256
`1ae4986ce53fab82afb4f7f4d8df50dbf04170fa81679e8ae259fd6e04ab4115`.
The tracked 136,357-byte source was never opened and remains SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`.
The separate isolated clear candidate opened without repair or conversion and
showed row 5 at its effective 20 pt `Default`, column C at its effective 98 pt
`Default`, and preserved explicit column F at 98 pt. Save/close/reopen retained
those values, `B2`/`B3`, the 22-by-7 topology, and headers 1/1/0. Its post-save
valid 43-entry ZIP is 136,938 bytes, inode 43005403, with SHA-256
`4db6e27e3893ca1892dd61dc221640d06c912db8f4f1d4aa51498607c5ca1325`.
The combined explicit and clear runs establish semantic native acceptance, not
byte-inertness or exact-artifact preservation: the explicit candidate changed
by 258 bytes, and earlier read-only evidence also showed silent normalization.
No general resize capability follows.

Focused dimension tests pass 13/13; codec tests pass 11/11; complete protos
passes 194/194 plus doctests. Numbers passes 241 library tests with four
ignored, 91 integration tests, and five doctests with one ignored. Archive
passes 130/130 plus doctests. Focused all-target check and strict all-target
Clippy pass. Boundary units pass 243/243 with zero retirement/API findings.
Host all-target/all-feature check/no-run, retained-axis 2/2, Pages-layout 1/1,
generated-roundtrip 1/1, scoped boundary 6/6, strict library Clippy,
formatting, and diff gates pass. Broad host all-target Clippy retains nine
unrelated existing lints.

## 2026-08-14 Numbers formula decoupling, not host exit

The focused Numbers facade now owns bounded semantic formula authoring and its
transaction, and the formula-specific boundary audit finds no focused import or
reexport of host formula identities. The root aliases previously exposed by
`litchi-numbers` are removed rather than redirected. Native and performance
evidence for the focused replacement is accepted in ADR 0008.

This cut deliberately does not declare a `litchi-iwa` host exit. The host still
contains production formula readers and setters, compatibility examples and
tests, raw formula vocabulary, and pivot-category behavior. Its Numbers formula
adapter imports the legacy `Formula*` vocabulary from neutral
`litchi-iwa-common`, preserving pivot compatibility without restoring
`PivotCategory` or raw identities in the focused API. No manifest dependency or
ordered debt entry is removed here.

A later host-exit cut must inventory and delete the remaining production
setters, raw exports, pivot vocabulary, dependent examples/tests, compiler and
mutation seams, manifest edge, and debt record together. It must rerun the
retired-surface and raw-identity audits and preserve all still-required reader
and cross-format behavior. Compatibility aliases, test-only reexports, or a
hidden facade bridge are not an acceptable substitute for that deletion.

## 2026-08-14 amendment: native-negative enforcement and lazy-owner hardening

The migration host again conforms to the accepted negative Keynote evidence.
`KeynoteEditor::set_slide_name`, its direct example, its README call, and the
two self-read-only mutation tests are deleted. The field-10 splice had returned
despite ADRs 0004 and 0008 recording that both focused and legacy prototypes
failed the real Keynote navigator gate. The existing declaration audit now
ratchets the setter, and additionally rejects restoration of the example or
README call. Semantic name reads and selector resolution remain supported;
name mutation remains blocked until the native-authoritative graph/cache
closure is discovered and passes an app-authored set/clear gate.

The concrete Keynote reader also stops eagerly materializing generated
`KN.PlaceholderArchive` and `KN.NoteArchive` values. Placeholder-owned and
speaker-note storage references now come from the existing strict, bounded
Buffa lazy projections, with byte/field/nesting/work failures mapped into the
format-owned semantic error model. The separate `TSWP.ShapeInfoArchive` path
remains generated-Prost debt because no equivalent strict projection owns it
yet. This is a two-owner decode cut, not a whole-graph Buffa or Keynote host
exit claim.

The focused Numbers formula transaction now recognizes an exact no-op when an
existing cached text is represented by a native string-list key. It resolves
that key against authoritative root and segmented string-list content through
a strict streaming visitor, charges the complete decode/list/search work, and
rejects duplicate keys or segment-range contradictions. Equal formula plus
equal resolved text shares the exact source artifact; different text remains a
real transactional rewrite. This is focused-owner correctness hardening only:
the host formula setters, compiler, pivot behavior, manifest edge, and ordered
debt remain exactly as recorded in the preceding amendment.

The root facade fuzz suite now registers `numbers_formula_cells`. Its bounded
native-seed command model exercises formula staging, changed/no-op commit,
exact apply/inverse, duplicate operations, cycles, cache mismatches,
foreign-source conflicts, constructor failures, redaction, and source
atomicity. This extends adversarial coverage without adding a new public owner
or changing the monolith-exit topology.

## 2026-08-14 amendment: Pages rooted-body lazy projection

The focused Pages rooted-body reader no longer decodes
`TSWP.StorageArchive` through Prost. The shared text-wire owner now exposes a
fully validated archive-free `ValidatedStorage`: it performs the existing
strict full known-tree preflight, enters exactly one bounded Buffa fragment
view, and fallibly materializes one semantic UTF-8 buffer plus checked runs.
Empty field-3 fragments remain zero-length runs, and UTF-8 length, UTF-16
length, fragment count, and semantic coverage must agree before publication.
Pages streams the field-17 section entries through the existing strict Buffa
boundary codec, retains only a bounded vector of semantic section references,
and projects names and section spans without a generated storage object. The
ordinary Pages dependency on `prost` moves to test-only oracle/fixture scope.

The cut is ratcheted at the production seam. Boundary policy rejects
`prost::Message`, direct generated `tswp` use, `StorageArchive::decode`, and
`litchi_iwa_text_wire::from_archive` in the production portion of the focused
Pages package, and rejects `prost` as a normal or target-normal Pages
dependency. Malformed known storage tables still fail before Buffa projection;
strict validation reasons remain observable through the established Pages
invalid-format envelope. Aggregate section-name bytes are now checked before
owned-string allocation. Concurrency coverage exercises independent
selector-first section-text commits, and the registered `pages_section_text`
sanitizer target covers commit/apply/inverse, selectors, spans, malformed
inputs, conflicts, redaction, and source atomicity.

This is a focused read-path and boundary-hardening cut, not a Pages host exit.
The migration host's body/section-text editor remains necessary for dependent
footnote and inline-object cleanup and for changed legacy nested packages that
the focused preservation-safe API intentionally refuses. Removing it now
would discard supported behavior. `PagesEditor`, creation, tables, media,
formatting, lifecycle work, ordered debt 017, and the
`litchi-iwa -> litchi-pages` edge therefore remain open.

The Keynote prepared-source coordination bridge now matches Pages and Numbers:
its hidden public entry points and their classification counter exist only
with `internal-iwork-source`. The boundary-policy snapshot also records eight
previously unclassified current manifest edges—two normal `xml-minifier`
edges and six dev-only `soapberry-zip` edges—without changing migration-debt
ordering. No hidden compatibility alias or new host dependency is introduced.

## 2026-08-15 amendment: focused Keynote production-Prost exit

The focused Keynote package no longer directly decodes application payloads
through generated Prost messages in production. The remaining eight eager
sites are removed together: slide node, slide, build, and shape-owner reads in
`package.rs`, plus the two color and two path-source validations in the slide-
transition transaction. Slide ownership and placeholder/note fields reuse the
strict bounded speaker-notes Buffa projection; transition semantics reuse the
strict bounded transition Buffa projection. Repeated build/drawable references,
build fallback durations, the slide-node edge, and the shape-owned-storage
edge are projected through complete bounded raw-wire preflights without
publishing generated types.

Opaque transition colors and paths now have strict schema-aware validation in
place of generated decode. Selected keys and scalar framing are canonical,
strings are UTF-8, Booleans are 0/1, floats are finite, singular fields are
unique, and the nested proto2-required point, size, element, connection,
subpath, and editable-node envelopes are proven before the bytes can enter a
semantic transaction. Unknown non-group fields remain caller-owned opaque
bytes. The ordinary `litchi-keynote` manifest edge to `prost` moves to
development-only oracle/fixture scope.

A permanent boundary ratchet scans every focused Keynote production source
prefix and rejects `prost::Message`, generated-message `decode`, the former
generic decode helper, and any restored normal/build Prost manifest edge. Its
258 regression cases and the live 64-package/239-declaration boundary graph
pass with the existing 14 ordered monolith debts unchanged. This does not
claim that transitive archive/proto internals are Prost-free, nor does it
retire any remaining `KeynoteEditor`, creation, media, chart, shape, table, or
layout owner. Ordered debt 014 and the `litchi-iwa -> litchi-keynote` edge stay
open.

The same integration cut closes an allocation-amplification gap in focused
Numbers name edits. An edit now checks the aggregate retained bytes for every
staged before/after name before allocating either `Arc<str>`; exact no-ops
share one `Arc`, while checked overflow and one-over-limit cases return the
typed redacted `NameBytes` error. This is transaction hardening only and does
not change Numbers host ownership or ordered debt 015.

The root fuzz suite registers `keynote_slide_deletion` with a bounded native
three-slide command seed. It covers selector staging, successful deletion,
exact apply/inverse, conflicts, final-slide refusal, semantic/input limits,
redaction, malformed ingress, and source atomicity. This adds adversarial
evidence for the already focused deletion owner; it does not broaden deletion
semantics or remove another host surface.

## 2026-08-16 amendment: Numbers package root/sheet/text projection

The focused Numbers package removes the remaining eager Prost decodes at its
root-and-text reader seam for this slice. The type-1 document order is now a
strict, bounded `numbers_sheet_order_codec` projection. Type-2 standard and
type-3 form-based sheets use the existing strict name/drawable preflight and
publish only a semantic name plus validated drawable references. Selected
rich-text storage messages use the shared archive-free `ValidatedStorage`
decoder from `litchi-iwa-text-wire`; the compatibility diagnostic preserves
its historical deterministic ordering and ignores malformed storage
candidates. No generated document, sheet, form-sheet, or storage object is
published by these paths.

The projection charges physical input, fields, nesting, and aggregate wire
work, then charges rooted references, drawable references, title/name bytes,
and aggregate text before owned allocation. Canonical framing, unique
ownership, missing references, UTF-8, strict/Buffa parity, and caller-selected
semantic ceilings are checked before the rooted document becomes visible.
Unknown component and protobuf bytes remain source-owned and are not retained
or re-encoded by the private views. The boundary is deliberately narrower
than the complete Numbers graph: table, tile, formula, sidecar, and host editor
paths remain ordered migration debt, and this amendment does not claim the
monolith is removable.

The corresponding native acceptance gate is read-only. A mode-0444 copy of
the tracked `test-data/iwork/numbers/basic.numbers` fixture is opened in Apple
Numbers 14.4 (7043.0.93), with no repair, recovery, conversion, or warning UI.
The locked copy must show `Sheet 1`/`Table 1`, a 22-by-7 table, the
`Litchi native Numbers fixture` marker, and numeric `42`; focused reread must
report the same sheet, table, text, and cell values. The tracked source remains
unopened and exact. This is application-acceptance evidence for the selected
reader path, not a native resave, rendering, performance, or complete Buffa
claim.

## 2026-08-16 amendment: Numbers names guard leaves generated-Prost ingress

The focused Numbers names owner removes its last production generated-message
reads from the changed-only dependency and pivot guards. Calculation-engine
and formula-owner dependency envelopes now use the existing bounded strict
Buffa projections, while the raw root/reference records remain the authority
for locality, metadata, duplicate detection, and exact preservation. The
volatile sheet/table-name dependency rule is unchanged. Pivot refusal uses a
narrow field-85 `TSP.Reference` preflight and does not force the full table
storage graph; its conservative rooted traversal is charged before native
work and still refuses every changed table rename with a pivot owner.

The names production-prefix boundary audit rejects `prost::Message`, generated
archive decode calls, and generic generated decode helpers, while test-only
fixture builders retain the canonical Prost oracle. Focused tests cover valid
and malformed dependency routes, duplicate optional fields, empty versus
nonempty volatile sets, pivot/lock/form/legacy behavior, strict wire and
unknown-field preservation, work-limit refusal, source atomicity, and the
existing selector-first transaction semantics. This is a production decode
boundary and boundedness improvement, not a host-method retirement: no
`litchi-iwa -> litchi-numbers` edge, ordered debt 015, or migration-host
dependency is removed here. The complete Numbers table, tile, formula,
sidecar, and editor graph remains outside this focused exit.

The previously accepted Numbers names native oracle remains authoritative for
semantic and application acceptance: the Rust Unicode candidate and its
Numbers 14.4 Save As/reopen artifact retain the recorded hashes and visible
sheet/table/data values, and the independent locked-table oracle retains its
negative behavior. Because those app-authored artifacts do not contain a
pivot or volatile-name dependency, native UI evidence does not replace the
synthetic malformed/guard matrix for the new refusal branches.

## 2026-08-17 amendment: Numbers rich-text envelope projection

The focused Numbers table extractor now consumes the strict raw-wire
projection of its type-6218 rich-text payload directly. The projection requires
one canonical local storage reference and one cell-owner field, validates the
nested reference, and passes the resulting storage identifier into the shared
bounded text-wire owner. The former generated Prost
`RichTextPayloadArchive::decode` was only used to recover and parity-check that
same identifier; it is removed from production, so the selected path no longer
allocates a complete generated envelope before its text/storage budgets are
applied.

This is a focused duplicate-parse and allocation cut. Unknown component and
protobuf bytes remain source-backed, and the existing malformed-candidate,
aggregate text, reference, field, nesting, wire-work, and failure-atomicity
contracts remain in force. It does not transfer the table model, tile/BNC
cells, list/formula/comment/AST graph, Numbers host editor, or any public raw
identifier surface. The `litchi-iwa -> litchi-numbers` edge and ordered debt
015 therefore remain open; all 14 ordered monolith debts are unchanged.

The focused all-feature Numbers suite currently passes 298 tests with four
ignored, including nine focused rich-text projection tests; document-reader
integration passes 16/16, and the boundary suite passes 267/267 while the live
graph remains 64 packages and 240 internal
declarations. Existing native Numbers basic and formula/rich-text fixtures
continue to provide application semantic acceptance; this read-only
implementation change introduces no native save/reopen or performance claim.

## 2026-08-20 amendment: selective iWork semantic package ingress (scoped follow-up)

This scoped follow-up refines and supersedes the broad ingress wording added
earlier on 2026-08-20. The following matrix is the normative boundary for
packaged and directory ingress; a semantic `Document` is never an exact-package
writer.

| Caller and path | Platform | Profile and selected members | Source lifetime and writer boundary |
| --- | --- | --- | --- |
| `litchi::iwork::Document::from_bytes` / `from_shared_bytes` | all | `SEMANTIC_METADATA`: canonical IWA members plus `Metadata/Properties.plist`, `Metadata/BuildVersionHistory.plist`, and `Metadata/DocumentIdentifier` | The `PreparedSource`/`SourceCatalog` retains the borrowed or shared source only through validation and handoff; the semantic Document drops it and has no writer |
| `litchi::iwork::Document::open` regular file | non-Windows | `SemanticMetadataAny` with the same three authorities after root inspection | The prepared package/catalog may retain exact bytes until semantic handoff; the published Document does not |
| `litchi::iwork::Document::open` regular file | Windows | `SemanticMetadataAny` falls back to the generic full `SourceCatalog` because the path adapter cannot provide the selective capture invariant | Media may be materialized during this fallback; the semantic Document still drops source bytes and has no writer |
| `litchi::iwork::Document::open` directory | all | Frozen directory semantic profile with the same three authorities | No exact ZIP allocation exists; the directory bundle is released before the semantic Document is published |
| `litchi_keynote::Document::open` regular file | non-Windows | `KeynoteProperties` / `SEMANTIC_PROPERTIES`: canonical IWA members plus only `Metadata/Properties.plist` | `PreparedSource`/`SourceCatalog` can preserve the exact source until the semantic handoff; the Keynote Document has no writer |
| `litchi_keynote::Document::open` regular file | Windows | The Keynote properties path retains its generic full-catalog fallback | Media may be materialized by the fallback; the published Keynote Document has no writer |
| `litchi_keynote::Document::open` directory | all | Existing frozen Properties-only directory profile | No exact ZIP source is retained and no writer is exposed |
| Full `SourceCatalog` / format-owned `litchi_keynote::Package` preserve path | all | Full catalog materialization, including untouched media and opaque members | `SourceCatalog::write_to`/`to_bytes` is the low-level exact byte-copy path while the catalog lives; the format-owned exact writer is limited to the full Package path |

The facade and Pages semantic profiles use the three bounded metadata
authorities; Keynote's archive-free profile intentionally selects only
Properties. Byte and shared-byte semantic preparation is selective on every
platform. Windows `SemanticMetadataAny` and the Windows Keynote properties
path are the documented full-catalog exceptions. Directory preparation is
selective but cannot claim exact ZIP source identity. Exact source retention
therefore belongs only to the live `PreparedSource`/`SourceCatalog` preserve
handoff; semantic Documents and directory-backed snapshots do not retain it.

Container-level ZIP checks remain in force. A selected logical profile applies
its 64 KiB and supported-compression checks before selected payload
materialization, with the existing typed archive `InvalidBundle`/`Limit`
categories and facade mapping. In the selective Keynote Properties profile,
malformed, opaque, or oversized BuildVersionHistory/DocumentIdentifier
payloads are outside the profile and are ignored; the same conditions on
selected authorities remain errors. Paired alias-only metadata names remain
ignored, one-sided exact/alias authority collisions are rejected, and
ordinary exact local/central mismatches retain their existing diagnostic.

### Evidence and pending performance-matrix entries

The completed evidence is archive-entry instrumentation only. It proves the
selection/read shape but is not latency, allocation, resident-memory, or
throughput evidence:

| Profile/path evidence | Selected reads | Full reads | Latency | Allocation/RSS |
| --- | ---: | ---: | --- | --- |
| Modern synthetic `SEMANTIC_METADATA` catalog (8 members) | 5 | 8 | `PENDING — no benchmark run` | `PENDING — no benchmark run` |
| Modern synthetic `SEMANTIC_PROPERTIES` catalog (7 members) | 2 | 7 | `PENDING — no benchmark run` | `PENDING — no benchmark run` |
| Shared `PreparedSource` with `SEMANTIC_METADATA` | 1 root inspection + 5 selected-catalog reads = 6 | — | `PENDING — no benchmark run` | `PENDING — no benchmark run` |
| Windows `SemanticMetadataAny` packaged-file fallback | `PENDING — full-catalog path` | `PENDING — no benchmark run` | `PENDING — no benchmark run` | `PENDING — no benchmark run` |

These are change-record placeholders, not fabricated performance results. A
future perf-baseline entry may replace the pending cells only after measuring
the same profiles with representative wall-clock, copied-byte/allocation,
and resident-memory evidence.

### Reconciliation with the earlier 2026-08-20 wording

The earlier statement that read-only iWork ingress and archive-free Keynote
used one three-authority profile is superseded by the matrix: generic facade
and Pages paths use three authorities, while Keynote is Properties-only. The
earlier statement that media and unknown members remain in a retained exact
source is narrowed to the live `PreparedSource`/`SourceCatalog` preserve
handoff; semantic Documents and directories do not retain a ZIP source. The
earlier exact-writer statement is narrowed to the low-level live
`SourceCatalog` byte-copy and the full format-owned Package path. Finally, the
earlier instrumentation counts are read-count evidence only; no latency or
memory improvement is claimed until the pending performance entries are
measured.

## 2026-08-20 amendment: Numbers type-6002 Tile production-decode exit

The focused Numbers type-6002 Tile owner now uses the no-`Vec` transactional
`numbers_table_cell_storage_codec::decode_tile_with_visitor` route. Its strict
handwritten canonical router validates root/row framing, required and duplicate
fields, canonical scalar and Boolean encodings, unknown/group forms, and
aggregate limits. Private Buffa lazy views are used only as a parity check after
strict traversal. Production publishes no generated `tst::Tile`; the permanent
boundary check permits Prost only in test fixtures and differential oracles and
rejects both generated `Tile::decode` and `TileRowInfo::decode` (including
`tst::Tile::decode` and `tst::TileRowInfo::decode`) in production. Row callbacks
borrow storage and offset slices from the source component.

The visitor guards the prefix materialized-cell total before parsing a row. A
semantic cell error is retained while later rows still receive strict wire
validation, producing the complete decode report without growing the semantic
table. The package merges that full report, then charges the aggregate
materialized-cell count, then returns the retained semantic error. A later
malformed wire error therefore wins over an earlier semantic cell error; an
aggregate cell limit wins over a retained semantic error. Candidate tables and
retained budgets are local and publish only after success, while bounded decode
work remains charged after rejection.

This turn also hardens legacy type-6000 admission: strict shape classification
still ignores TableInfo false positives, while schema-shaped legacy models hit
the table budget before Prost decode/materialization and cannot commit rejected
candidate budgets.

The production cut removes the duplicate Tile parse and row-buffer copies, but
source borrowing at this codec boundary is not an end-to-end zero-copy claim.
The 2026-08-20 gates record 221 `litchi-iwa-protos` tests, 311 passed and four
ignored Numbers unit tests, and all 14 integration binaries passed (120 tests
total), including `document_reader` 16/16 and `tile_reader_integration` 2/2.
Boundary is 271; the live graph is 64 packages, 240 internal declarations, and
14 unchanged ordered debts. The final temp-corpus nightly ASan smoke ran 100
runs with eight seeds; the read-only native Numbers ZIP/directory hash and UI
gate passed its known hashes.

This is a focused production-decode exit, not completion of the extractor,
Numbers package, or monolith. No measured performance, latency, RSS, native
save, or native mutation claim follows. Eight production generated decodes
remain in the extractor, old `litchi-iwa` Tile paths remain, and debt 015 plus
all 14 ordered migration debts are unchanged.

## 2026-08-21 amendment: Numbers TableDataList/Segment bounded ingress

The focused Numbers package now routes `TST.TableDataList` and
`TST.TableDataListSegment` through a strict handwritten wire visitor backed by
a private Buffa lazy projection. Package and Document use the same strict
ingress. Document skips comment resolution by design; Package remains strict
when retaining comments. Candidate publication retains only final semantic
table values and bounded segment-id/key state, with fallible reserves; this is
not a claim that all intermediate state is absent or that the path is
end-to-end zero-copy.

Required/duplicate/canonical-wire/UTF-8/reference/range/ownership checks,
candidate atomicity, and later-wire-error precedence remain enforced. The
package ceilings are 512 MiB input and output, 1,000,000 fields, nesting 64,
and 16,000,000 work units, with tighter payload and text budgets before
allocation. Five generated extractor decodes remain (four table-model and one
FormulaArchive), while FormulaArchive/legacy `litchi-iwa` comment migration
debt, legacy monolith/dependency edges, and all 14 ordered debts remain open.
The former generated comment decode is no longer part of this extractor path.

Focused evidence on 2026-08-21 is 237 protocol tests, 49 extractor tests, 14
integration tests, 281 boundary tests, and an 11-seed fuzz corpus; the TDL
nightly ASan smoke completed 100 runs without an artifact. A disposable native
1,200-by-8 Numbers document was saved and reopened with strings, formulas, an
intentional formula error, rich text, and a persisted comment and showed no
repair dialog. The archive has DataList members, but no parsed type-6011
Segment evidence was established. This is a correctness/boundedness slice
only: no native segment or save-mutation, performance, host-exit, or complete
monolith-retirement claim is made.

## 2026-08-21 amendment: Numbers CommentStorage bounded migration

The focused Numbers package now routes `TSD.CommentStorageArchive` through a
strict handwritten CommentStorage codec backed by a private Buffa lazy
projection. Package remains strict for comment validation and resolution;
Document deliberately skips comment resolution. Per-cell comment
materialization is bounded and fallible, with strict validation before
candidate-local publication and fallible reserves for the materialized value.
The private projection is parity-only; raw source bytes remain authoritative.

FormulaArchive and the legacy `litchi-iwa` comment path remain migration debt.
Five generated extractor decodes remain (four table-model decodes and one
FormulaArchive decode), and the former generated CommentStorage decode is no
longer part of the production extractor path. Legacy monolith/dependency edges
and all 14 ordered debts remain open.

Focused evidence on 2026-08-21 is 237 protocol tests, 324 passing Numbers
library tests with four ignored (including 53 focused extractor tests), 14
integration tests, and 281 boundary tests. The CommentStorage fuzz corpus has
20 seeds, and its nightly ASan smoke completed 100 runs cleanly without an
artifact. A disposable native 1,200-by-8 Numbers document with comment,
formula, and rich-text content was saved and reopened without a repair dialog.
The archive has DataList members, but no parsed type-6011 Segment proof was
established. This is a bounded correctness slice only: no native segment or
save-mutation, performance, host-exit, or complete monolith-retirement claim
is made.

## 2026-08-21 amendment: focused detector host-edge retirement

The legacy `litchi-iwa` host no longer depends on `litchi-iwa-detect`. Its
public detector facade, detector-only bundle helper, host detector assertions,
and detector fuzz import were removed. The unified host reader, comment
package-root validation, and Numbers root selection now share a small private
canonical `DocumentArchive` classifier. It recognizes the Pages, Numbers, and
Keynote root shapes without consulting raw message-type IDs or publishing a
detector vocabulary.

The focused detector remains the sole owner of packaged-file, seekable-reader,
and legacy-directory ingress. Its malformed/ambiguous-root, marker-conflict,
and cursor-restoration coverage is unchanged; generated host roundtrip
coverage still exercises all three unified and focused readers. This closes
ordered debt 006 and removes exactly one normal host dependency declaration.
No other host migration edge or semantic ownership claim changes here.

The current topology is 64 workspace packages, 239 internal dependency
declarations, and 13 ordered migration debts. The remaining host readers,
editors, examples/tests/fuzz ownership, and the broader monolith exit remain
open.

## 2026-08-22 prior-commit section-text bridge and aggregate gate evidence (historical; predates current-wave amendments)

The prior committed record verifies the Pages host section-text selector-first
bridge. This is historical evidence, not a current-wave gate. The
bridge rejects U+FFFC, uses strict invalid-format fallback, and passes the
topology, readback, and budget proofs. The Numbers scalar shortcut
interop/no-op test passes and preserves the source on the no-op path. The XLS
gate records 1,274 passed; the boundary, iWork-leaf, and root checks record
301, 13, and 10 respectively.

The prior all-target aggregate gates were green: Pages all-targets include the
reported 84+15+14+1+10+8+5+6+7+4 runs; Numbers records 374 library tests
plus green integration targets; `litchi-iwa` records 1,520 library tests; and
strict Clippy passes. Native evidence includes the Pages lock reopen and a
native Numbers 1,200-by-8 formula/error/rich-text/comment reopen. Direct
type-6011 proof remains explicitly unconfirmed.

This is bounded verification evidence only. No host dependency edge or
ordered debt is retired, and this wave makes no full monolith-exit or
complete Prost-migration claim.

## 2026-08-22 amendment: focused semantic read and bounded API gates

The Pages package now owns the immutable semantic read of body footnotes
through `Package::body_footnotes()`. Its bounded strict projection validates
the body boundary, source-order UTF-16 positions, anchors, references, and
marker identity before publishing only public footnote values; native object
IDs and storage/marker attachments remain private. The focused package test
reads the projection repeatedly and verifies that the source bytes are
unchanged. The Pages host editor still owns footnote CRUD, so this read-owner
slice does not retire the `litchi-iwa -> litchi-pages` edge or its ordered
debt, and it makes no native-mutation determination.

The accepted report records a native Pages body/section/footnote fixture
reopening without a repair, recovery, conversion, or warning UI. The reported
artifact is 100,389 bytes with SHA-256
`0ac8f1db0fb64556031a05284f7e7e7ec11d70a7dcbccd0a2c698365fd355fc2`; the
artifact is absent from this checkout, so this remains reported/accepted
historical evidence only. Its reported pre-close and reopened accessibility
snapshots match while retaining the body, section, and footnote text. This is
native semantic no-repair evidence for the inspected fixture only; it is not
byte-inert-save evidence, a general native-mutation claim, or host-edge
retirement.

The focused Numbers API gate verifies name, index, and typed `Position`
shortcuts for sheet and table selectors, semantic `Package::table` lookup,
and a presence-preserving A1 view that distinguishes a stored empty cell from
a missing cell. Scalar set/clear shortcuts reuse the existing source-bound
atomic transaction and inverse path; the failure cases leave the source
unchanged. The typed index gate verifies `Reference` iteration through
`ObjectIndex::references()`: ordinary sources remain deterministic and
sorted, while an explicit per-source opt-in preserves insertion order. These
are selector and index API gates only; no host edge or ordered debt changes.

The transition gate verifies bounded opaque payloads are fallibly reserved and
copied before publication, maps allocation failure to the existing typed
`PayloadTooLarge` error, and retains the copied semantic payload after its
source is changed. This is allocation/ownership hardening only; no native
mutation conclusion follows from it. The deleted
`crates/litchi-iwa/examples/edit_pages_section_text.rs` was only the duplicate
migration-host example; the canonical
`crates/litchi-pages/examples/edit_section_text.rs` remains. `PagesEditor`,
the other host examples, and the host dependency edge remain migration
surface.

These focused Pages, Numbers, index, and transition gates, together with the
example-inventory change, are verified correctness and ownership evidence for
this amendment. They do not close the remaining host edges or ordered debts,
and the monolith exit remains incomplete.

## 2026-08-22 pre-gate worktree evidence disposition

The current working wave adds further strict codec and owner seams, including
Pages footnote/movie-caption inputs, Numbers table-model inputs, Keynote chart
title inputs, archive entry views, source-backed text handling, and the
archive-free transition projection. These changes are not admitted as a new
deletion-gate result yet. The current `litchi-iwa-protos` `cargo test` reaches
339 tests: 338 pass and one known strict-model failure remains. The protocol
gate therefore remains unadmitted for this wave, and this result is recorded
for diagnosis only rather than as a current gate.

The available current-wave aggregate runs are also not clean: the recorded
`litchi-iwa` library run ended at 1,231 passed and 274 failed, and the Numbers
library run ended at 360 passed, 7 failed, and 4 ignored with integration
targets also failing.
Those results cannot satisfy deletion gates 3 or 4 and cannot support a host
edge-retirement claim. The prior clean counts above remain historical rather
than evidence for these changes. The new fuzz manifests and corpora document
compile-check commands, but no completed current-wave sanitizer smoke result
is accepted; the existing dated fuzz gates remain the only recorded fuzz
evidence. The separate native Pages body/footnote run used a 100,426-byte
artifact with SHA-256
`bfdd06db92af66df5cb434858ff8274eba9f1a8fe2135a3e711526e28edf504b`; its
pre-close and reopened accessibility snapshots compare equal and retain the
body, section, and footnote text without a repair or conversion prompt. A
separate Keynote chart-title cycle saved and reopened `Revenue by region`,
then cleared the visible title (two AX matches before/after set, zero after
clear, and the chart-title checkbox false); both 57-file ZIP checks passed and
no repair or warning UI appeared. The source, set, and cleared artifacts have
SHA-256 values `ee57fa54e90a8256b5b9973c24233ac4a2ec8b506614fc725e048c4195f`,
`473c8df61e1e118e1768926ce4dd8eaf05096651a56546e04ea7d1967cfd6a1d`, and
`c728fd9b1c5283fd107bd2efe330d1a320a565d84442075a4091bd3c69306a13`.
These are standalone native semantic
no-repair results for the inspected artifacts only, not Litchi save-mutation,
byte-inertness, or deletion-gate evidence while the source/build and
aggregate test gates are incomplete.

This disposition records no new debt retirement, dependency-edge change,
native mutation/save claim, performance result, or monolith-exit progress.

## 2026-08-22/23 amendment: committed bounded-projection wave verification

This amendment preserves prior committed-baseline evidence, but the cited
detached commit/tree is unavailable in this checkout, so its provenance is
unavailable. The counts below are not current-wave gates. The current wave
remains pending its final checks; its worktree changes are not verification
evidence and are not promoted by this record.

The prior committed baseline reported a clean focused protocol-codec check
with 33 passed tests and a full `litchi-iwa-protos` suite with 267 passed and 0
failed; its Pages all-target gate recorded 149 passed; its Numbers library
gate recorded 356 passed with four ignored; its Keynote all-target gate passed;
and its `litchi-iwa` library gate recorded 1,442 passed. The prior committed
root `litchi` facade/API and rustdoc checks were free of a public
migration-host, generated-message, native-ID, or BNC bridge. Its boundary
check reported 64 workspace packages, 239 internal dependency declarations,
and exactly 13 ordered migration debts. None of these baseline counts is a
current-wave gate.

The prior baseline's native evidence remains semantic no-repair evidence for
inspected artifacts only. It includes the accepted Keynote transition and
title and Pages body/background reopen paths, plus the Numbers
rich/formula/comment reopen that retained `Persistent TDL comment` and whose
archive inspection found 24 type-6005 entries and 0 type-6011 entries. The
separate recorded Pages body/footnote artifact is 100,426 bytes with SHA-256
`bfdd06db92af66df5cb434858ff8274eba9f1a8fe2135a3e711526e28edf504b`; its
pre-close and reopened accessibility snapshots agree. The separate Keynote
chart-title cycle reopened `Revenue by region`, cleared the visible title,
passed both 57-file ZIP checks, and has source/set/cleared hashes
`ee57fa54e90a8256b5b9973c24233ac4a2ec8b506614fc725e048c4195f`,
`473c8df61e1e118e1768926ce4dd8eaf05096651a56546e04ea7d1967cfd6a1d`, and
`c728fd9b1c5283fd107bd2efe330d1a320a565d84442075a4091bd3c69306a13`.
These baseline native runs do not prove Litchi save mutation, byte-inertness,
native segment support, performance, or a current-wave deletion-gate result.

The baseline fuzz manifests and checked-in corpora establish target shape and
compile/list coverage only. The dated TableDataList and CommentStorage ASan
smokes remain the recorded baseline sanitizer evidence (100 runs with 11 and
20 seeds respectively). No current-wave sanitizer campaign or current-wave
fuzz gate is claimed, and a stable fuzz build or `cargo check` is not
sanitizer-backed fuzz evidence.

The current wave retires no host dependency edge and removes no debt pending
final checks. The authoritative baseline ledger remains at 13 ordered debts,
including the remaining `litchi-iwa` format-owner edges; host
readers/editors, examples, tests, and fuzz ownership remain migration work.
The monolith deletion gate therefore remains open.

## Current present status

The recorded detached baseline (its commit/tree provenance is unavailable in
this checkout) records 64 workspace packages,
239 internal dependency declarations, and 13 ordered migration debts. Its
focused protocol codec check records 33 passed tests; the full
`litchi-iwa-protos` suite records 267 passed and 0 failed. The baseline
Numbers library gate records 356 passed tests with four ignored; the
`litchi-iwa` library gate records 1,442 passed; the Pages all-target gate
records 149 passed; and the Keynote all-target gate passes. The current wave
has no admitted gate result yet. The prior XLS certification gate remains at
2,146 passed across BIFF15/CFB254/XLSB603/XLS1274.

Native no-repair reopen evidence covers Keynote transition/title and Pages
body/background paths. A native Numbers rich/formula/comment reopen retained
the exact `Persistent TDL comment`; its archive inspection recorded 24
type-6005 entries and 0 type-6011 entries. This does not establish the
type-6011/Segment path, whose proof remains withheld.

FormulaArchive and the legacy `litchi-iwa` comment path remain migration debt;
legacy monolith/dependency edges and all 13 ordered debts remain open. These
checks provide focused correctness/boundedness and native no-repair evidence
only: no native segment/save-mutation, performance, host-exit, or complete
monolith-exit claim follows. The dated verification gates above remain
historical evidence, and the monolith exit remains incomplete.

## 2026-08-23 amendment: wave17 focused gates and native no-repair evidence

Wave17 records the Buffa guard at 12/12, the Pages footnote codec at 11/11,
and a successful associated fuzz-target check. Numbers TableDataList storage
records 37/37. The reviewer’s post-strict-model-fix report recorded the
`litchi-iwa-protos` result as 341/341 during wave17; this remains dated wave17
evidence, not a current-worktree result. This does not rewrite the historical
267/0 baseline or the earlier 339/338+1 pre-gate observation.

The supplied application record reports no repair for the named fixtures. The
Keynote focused report records 20+5 cases; its fixture is 503,480 bytes with SHA-256
`1d67d4263851487bc3868342884f77f46fa643e542e661951d222411c11d21d1`.
The Pages fixture is 104,319 bytes with SHA-256
`7e01c1caa6fe3b0f0b699d5a022c0f2daf3667a22896a53e65289f5297d3ddb2`; the
Numbers fixture is 136,527 bytes with SHA-256
`fa87cffc0669af1ca0ecc6c4e41932aee4619929fa303e52f06ea77887990d98`.
Numbers archive inspection found 24 type-6005 entries and 0 type-6011
entries. The zero type-6011 result is explicitly a no-claim: it does not
prove the native Segment path.

The Numbers formula-category change recorded for this wave is preflight
arithmetic only; category traversal and retained projection semantics are
unchanged. Pages footnote evidence keeps the per-projection semantic budget
separate from `TransactionBudget`; it does not account for staged `set` text,
`after` values, or candidate-package retention.

These dated results do not establish byte-inert save behavior, native Segment
support, or type-6011 proof. They retire no debt or host dependency edge and
make no monolith-deletion claim; the historical counts and authoritative
13-entry ledger remain unchanged.

## 2026-08-23 amendment: Keynote text allocation shape and soundtrack-order scope

The new `Slide::plain_text` path computes a checked UTF-8 output length for
title, content, non-empty text storages, and notes, reserves one destination
`String` when that length is representable, and appends values in the same
semantic order as `all_text().join("\n")`. Its focused cases cover empty
values and empty storages. This is structural allocation-shape and semantic
ordering evidence only; it is not an allocation-count, latency, throughput,
RSS, or peak-memory measurement.

The focused soundtrack-order transaction changes only the order of existing
soundtrack media references. Its package-level cases cover a synthetic bounded
media closure, exact-source no-op and inverse behavior, unknown-field and
untouched-entry preservation, malformed ownership rejection, and aggregate
reference limits. This scope does not claim native Keynote open/save evidence,
soundtrack settings or media CRUD ownership, host-edge retirement, or any
monolith-exit result. The monolith deletion gate therefore remains open.

## 2026-08-23 superseding amendment: Keynote soundtrack-reference order ownership

The preceding soundtrack-order paragraph is superseded only as to ownership;
its dated evidence remains historical. The concrete Keynote package now owns
the bounded `soundtrack::order` move transaction exposed by
`Package::{edit_soundtrack_order, apply_soundtrack_order}`. It validates the
rooted Document -> Show -> type-21 Soundtrack chain, streams and bounds the
existing field-3 media-reference sequence, and rewrites only its order while
preserving media assets, metadata, unknown fields, and unrelated components.

This is not soundtrack media/settings CRUD. `litchi-iwa` remains the owner of
the remaining media/settings CRUD and host compatibility, including item and
asset creation, add/insert/replace/remove/lifecycle operations, and retained
host readers/editors outside the focused package seams. The separate focused
`soundtrack::{Mode, Settings}` API remains the owner for playback mode and
volume.

The order transaction's synthetic bounded cases are not a native Keynote
open/save gate. No host dependency edge or ordered debt is retired, and this
amendment makes no host-exit, host-deletion, or complete monolith-exit claim;
the monolith deletion gate remains open.

## 2026-08-23 amendment: bounded Pages fallback/removal, Keynote chart-title bridge, and IWA entry delegation

The detached wave ref `696be8e2e` contains four source-level changes relevant
to the ordered exit: `8e7fba67b` narrows optional Pages footnote fallback,
`696be8e2e` makes legacy footnote removal failure-atomic,
`966fc0212` delegates legacy Keynote chart-title operations through the
focused semantic package, and `0b118f2d8` delegates IWA entry lookup to
`PackageState`. This amendment records bounded implementation scope only; it
does not add Cargo or native gate evidence.

The Pages fallback now recognizes only the known malformed optional-footnote
storage field-16 error forms, while unrelated invalid body payloads remain
strict. Legacy footnote removal edits a cloned editor, validates absence of
the removed native reference identity in the staged graph, and publishes only
after that postcondition succeeds. Thus adjacent-footnote position shifts do
not delete the wrong reference, and a failed postcondition does not partially
mutate the live editor. The focused Pages package, native application
behavior, and broader Pages host reader/editor ownership remain open.

The legacy Keynote chart-title read/set/clear and catalog routes now enter the
focused `litchi-keynote::Package` semantic chart-title API after the host
resolves and validates native chart ownership. Identifier-based calls map to
semantic chart positions; focused commits are written back to the host
editor. The bridge covers this chart-title seam only. Remaining chart graph
work, other chart mutations, and the host compatibility surface remain
migration work, and no native open/save or byte-inertness result is implied.

IWA package entry containment, reads, mutation helpers, replacement,
removal, and IWA insertion now use the package state's canonical position
lookup directly. This removes a duplicate helper but does not remove the
`litchi-iwa` host, its dependency edge, or any ordered debt.

The historical 64 workspace packages, 239 internal dependency declarations,
and 13-entry ordered migration ledger are unchanged. None of these bounded
changes satisfies the deletion gate, retires a host edge, or authorizes
monolith deletion; no Cargo/native/complete-monolith result is claimed.

## 2026-08-23 amendment: detached bounded ingress and surface hardening

The detached commits `a4e7066ea`, `e834a6ced`, `5be5af395`, and `fa2ad9cf2`,
together with corrective commit `b72e7e50f`, are recorded here only as
bounded hardening inputs to the existing migration work. Their exact source
paths are, respectively,
`crates/litchi-numbers/src/package/extractor.rs`,
`crates/litchi-iwa/src/pages/editor/footnotes.rs`,
`crates/litchi-iwa-protos/build.rs`, and
`crates/litchi-iwa/src/identity.rs` plus `crates/litchi-iwa/src/lib.rs`.

The Numbers visitor now bounds reply identity collection, rejects a direct
self-reference and duplicate reply IDs, and verifies visitor cardinality
against the codec report before publishing a compatibility candidate. Pages
footnote cleanup bounds body collections, uses checked counts and fallible
reservations for removed graph identifiers, and retains staged publication
with identity-based deletion. Corrective commit `b72e7e50f` changes the
cleanup helper to `Result<[u64; 3]>`; the `e834a6ced` body change alone is not
presented as compilable. The Pages provenance guard inventories the
additional live route consumers, requires private or crate-visible numeric
declarations, checks production decode/remap/write markers, and extends the
canonical document-field ratchet. The IWA document identity type and its
generation/access/regeneration methods are crate-private and no longer
re-exported from `litchi-iwa`.

These changes tighten malformed-graph handling, allocation bounds, route
provenance, and public-surface privacy; they do not move format ownership or
mutation responsibility. No Cargo/build command or native application run
was performed or admitted. This amendment makes no host-edge, ordered-debt,
or monolith-exit claim and does not alter the deletion gate.

## 2026-08-23 amendment: detached Keynote chart-title owner-scan hardening

Detached commit `612bc0f26` is limited to the focused Keynote package's
chart-title graph builder. With mutation guards enabled, it lazily constructs
one package-wide `HashMap` of non-style owner counts and shares that map across
all charts in the graph build. `ChartGraphScanBudget` accumulates component,
object, message, payload, and parsed-field work once against the package-wide
`WireWork` limit, replacing the prior repeated per-chart ownership scan while
retaining the exactly-one-owner and nonzero-chart checks.

This is package-level work hardening only. The aggregate `WireWork` accounting
covers the owner scan and nested chart-reference parsing; it is not a bound on
whole-chart CPU or latency, and no such result is claimed. No Cargo/build/test
or native result is added. This amendment makes no dependency-edge,
ordered-debt, migration-host, or monolith-exit claim, and it does not alter the
deletion gate.

## 2026-08-23 amendment: detached Pages output arithmetic and Numbers table-count bounds

Detached commit `79b2a3267` changes only
`crates/litchi-pages/src/package/footnote_text.rs`. Custom-mark rewrite output
length now uses checked additions for the encoded field tag, length varint, and
payload, propagating overflow as the existing typed `OutputBytes` limit error
before output-capacity arithmetic. The existing output-byte, field-count, and
rewrite-work checks remain the governing bounds; this records source-level
arithmetic hardening only.

Detached commit `5b82d9167` changes only
`crates/litchi-numbers/src/package/extractor.rs`. `TableDataExtractor` retains
the configured `max_tables` and clamps caller-provided limits to it. Before
each fallible one-slot result reservation, a checked next-count supplies the
allocation-error amount; the reservation itself does not enforce the table
count. Candidate admission and table-count increments use checked counting,
while checked `next_seen` likewise supplies the allocation-error amount for the
fallible seen-object reservation rather than enforcing a count. Over-limit
candidates retain the structured-table limit error and host-count overflow
remains an invalid-format error. The compatibility table fallback is therefore
bounded at admission and result collection without changing its ownership or
projection scope.

These two detached diffs are recorded as bounded source changes only. No
Cargo, native, dependency-edge, ordered-debt, host-exit, or monolith result or
claim is added by this amendment.

## 2026-08-23 amendment: Pages route provenance inventory completion

Commit `ebf7a9aa7` changes only `crates/litchi-iwa-protos/build.rs`. Its Pages
route inventory now covers the body-shape caption path and focused section
background, pagination, settings, text, and transaction consumers, in
addition to the existing legacy and footnote paths. Declaration paths and
production use-marker paths are tracked separately, and the registry check
retains the existing `2001u32` and `2022u32` storage dispatches. These are
build-time provenance checks; they do not publish new route IDs or API
surface.

The route source list adds `cargo:rerun-if-changed` coverage for present
workspace siblings. The guard remains all-or-none once any route source is
seen: a partial sibling checkout fails closed, while a completely absent
sibling inventory keeps the standalone-package path intentional. Numeric route
declarations may be private or restricted (`pub(crate)`, `pub(super)`, or
`pub(in crate::pages)`), but not unrestricted `pub`. No unavailable IDs are
added to canonical ratchets; the existing canonical message blocks remain the
authority.

This amendment records source-level guard hardening only. No Cargo/build/test
or native result is claimed, and no `litchi-iwa` dependency edge or ordered
debt is retired. It does not move ownership or satisfy any host-exit or
monolith-deletion gate.

## 2026-08-23 amendment: legacy Pages body-footnote borrowed read projection

Commit `588abfb6d` is a bounded change to the legacy Pages footnote path in
`crates/litchi-iwa/src/pages/editor/footnotes.rs`. A single
`FootnoteGraphBudget` now accounts across one rooted body-footnote traversal:
the body table, reference objects, storage payloads, marker tables, and the
selected nested messages. `FootnoteStorageProjection<'source>` borrows the
storage wire from package-owned bytes for reads and validation rather than
constructing a generated `tswp::StorageArchive`. It canonical-checks only the
selected storage kind, UTF-8 text, table framing/multiplicity, marker
index/reference, and bounded footnote-entry/reference fields needed for the
semantic footnote; the strict reference/body/marker codecs remain bounded
validators for their selected messages.

The budget applies checked cumulative input-byte, field, rewrite-work,
nesting, and allocation charges, derives strict-codec options from wire
preflight reports, maps preflight failures to typed limit/allocation errors,
and uses fallible reservations for projected text, entries, reference sets,
graphs, and owned strings. Body/table entries remain capped at 4096, and the
mutation-side encoded table rewrite has its own checked byte ceiling.

Generated Prost storage is confined to the mutation/template compatibility
route. `decode_storage_for_mutation` preflights before decoding the temporary
value; source-preserving wire rewrites retain unrelated fields and existing
raw entry payloads, while the read projection retains unknown storage bytes in
the original source instead of rebuilding them.

The scope is deliberately narrower than the host deletion gate. The change
does not provide a package-wide alias/ownership census: duplicate references
within the selected body are rejected, but global non-aliasing is not proved.
It does not claim canonicality for opaque nested unknown fields in table or
attachment payloads, nor full aggregate safety for every package graph,
archive, or mutation route. This records legacy-path bounded hardening only;
no Cargo/build/test or native result is admitted, no host dependency edge or
ordered debt is retired, and no migration-host or monolith-exit claim follows.

## 2026-08-23 amendment: focused Numbers comment ownership census

Commit `96539bf47` (`fix(numbers): prove global comment ownership`) is a
narrow hardening slice in the concrete Numbers package, not a broad comment
CRUD migration. `litchi-numbers::Package` now performs a package-wide census
before publishing a changed existing comment text: it walks comment-list root
messages (`6005`/`6201`), segment messages (`6011`), comment-storage payloads
(`3056`), and all table-model cell stores that can carry comment keys. The
census keeps bounded list, segment, table-reference, cell-key, storage,
author, reply, and UUID facts and proves the selected root entry/storage/key
chain is globally unique.

The focused owner rejects zero or aliased IDs, duplicate list/storage routes,
missing or duplicated selected entries, zero keys/refcounts/storage references,
out-of-range or overflowing segment key envelopes, and incompatible archive
message metadata. Payload references must be nonzero, unique, and consistent
with the metadata; storage authors, replies, reply counts, storage IDs, UUIDs,
table references, and cell keys are checked for self-edges, aliases,
duplicates, and required cross-links before mutation.

The only changed write admitted by this seam is text replacement on an
existing root-list comment with refcount one, one global storage occurrence,
and no replies. New comments and changed clears remain refused because the
owner graph would need to be rewired and cleaned atomically. Segment entries
are still readable after strict half-open range validation, but segment set
and clear return `UnsupportedDependency` without changing source bytes. The
updated selector-first integration case preserves root replacement/inverse
coverage and asserts segment set/clear refusal plus semantic reread/reopen.

The census and all refusal checks precede reassembly. A permitted replacement
is built into a separate candidate, reopened and semantically verified, and
returned as a reversible exact-source patch with source/target ownership and
cell-byte checks. This documents focused Numbers package behavior only: no
broad comment CRUD, Cargo/build/test result, native application result,
dependency-edge or ordered-debt retirement, migration-host exit, or monolith
deletion claim follows.

## 2026-08-23 amendment: legacy Pages body-footnote borrowed read projection

Commit `588abfb6d` is a bounded change to the legacy Pages footnote path in
`crates/litchi-iwa/src/pages/editor/footnotes.rs`. A single
`FootnoteGraphBudget` now accounts across one rooted body-footnote traversal:
the body table, reference objects, storage payloads, marker tables, and the
selected nested messages. `FootnoteStorageProjection<'source>` borrows the
storage wire from package-owned bytes for reads and validation rather than
constructing a generated `tswp::StorageArchive`. It canonical-checks only the
selected storage kind, UTF-8 text, table framing/multiplicity, marker
index/reference, and bounded footnote-entry/reference fields needed for the
semantic footnote; the strict reference/body/marker codecs remain bounded
validators for their selected messages.

The budget applies checked cumulative input-byte, field, rewrite-work,
nesting, and allocation charges, derives strict-codec options from wire
preflight reports, maps preflight failures to typed limit/allocation errors,
and uses fallible reservations for projected text, entries, reference sets,
graphs, and owned strings. Body/table entries remain capped at 4096, and the
mutation-side encoded table rewrite has its own checked byte ceiling.

Generated Prost storage is confined to the mutation/template compatibility
route. `decode_storage_for_mutation` preflights before decoding the temporary
value; source-preserving wire rewrites retain unrelated fields and existing
raw entry payloads, while the read projection retains unknown storage bytes in
the original source instead of rebuilding them.

The scope is deliberately narrower than the host deletion gate. The change
does not provide a package-wide alias/ownership census: duplicate references
within the selected body are rejected, but global non-aliasing is not proved.
It does not claim canonicality for opaque nested unknown fields in table or
attachment payloads, nor full aggregate safety for every package graph,
archive, or mutation route. This records legacy-path bounded hardening only;
no Cargo/build/test or native result is admitted, no host dependency edge or
ordered debt is retired, and no migration-host or monolith-exit claim follows.

## 2026-08-23 amendment: checked Numbers formula-work products and owner preflight

Detached commit `711e2b517` is limited to checked work accounting in
`crates/litchi-numbers/src/package/extractor.rs`. The changed
formula-reference, TableInfo, and table-model-name products now use checked
source-length multiplication before work charges; no formula-category scan
change is claimed. The owner preflight uses checked additions for source,
field, nested UUID, local-reference, and malformed-field work.

An owner accumulator overflow is reported as `LimitKind::RewriteWork` with
the hard-coded `MAX_FORMULA_WORK`, then mapped to the
`SemanticLimitKind::FormulaWork` error. A FormulaWork limit propagates
instead of being treated as an ordinary malformed compatibility candidate;
other malformed owner candidates retain their skip path.

This does not move FormulaArchive or formula-owner responsibility out of the
current package/host boundary and does not broaden formula authoring or
rendering. It is source-level finite-accounting hardening only: no
Cargo/build/test or native result is admitted, no `litchi-iwa` dependency edge
or ordered debt is retired, and no host-exit or monolith-deletion claim
follows.

## 2026-08-23 amendment: guarded legacy drawable-comment text ownership

Commit `0fb5175bd` is limited to `crates/litchi-iwa/src/comments.rs`.
Before an existing drawable-comment text change can update shared storage
in place, the host validates the direct reply graph and requires one direct
drawable user. It then removes the selected edge only on a private package
clone and runs the package-wide reference census across comment storage and
reply edges, package-metadata maps, external/data references, object
registries, and ambiguous identifiers. Duplicate selected metadata
references fail closed; any remaining reference keeps the existing
copy-on-write path.

This is a narrow legacy-host compatibility guard, not broad comment CRUD or
a replacement owner. No Cargo/build/test or native result is admitted; no
`litchi-iwa` dependency edge or ordered debt is retired, no host-exit claim
follows, and the monolith deletion gate remains open.

## 2026-08-23 amendment: focused Pages aggregate semantic footnote budget

Commit `58f824eb6` changes only
`crates/litchi-pages/src/package.rs` and
`crates/litchi-pages/src/package/footnote_text.rs`. `FootnoteSemanticBudget`
is separate for each `project_body_footnotes` projection and each
`native_footnotes` pass, rather than a shared package-wide accumulator. It
checks projected footnote text plus custom-marker bytes before per-value
owned strings are retained. The existing `MAX_BODY_FOOTNOTES` bound,
fallible collection reservations, and per-value text/custom-marker checks
remain governing limits.

Section transaction paths use `TransactionBudget` for transaction/candidate
work; this footnote-text transaction/rewrite seam uses its own wire/text
limits and retains per-value checks, including `Footnote::with_custom_mark`.
`FootnoteSemanticBudget` does not
account for staged `set` text, `after` values, or candidate-package
retention; no aggregate set/after_text/candidate-retention claim is made.

This is focused aggregate semantic-cap hardening only. No Cargo/build/test
or native result is admitted and no performance, allocation-count, RSS, or
throughput claim is made. No `litchi-iwa` dependency edge or ordered debt
is retired, no host-exit claim follows, and the monolith deletion gate
remains open.

## 2026-08-23 amendment: Pages body-table lock ownership proof hardening

Commit `2219b53b4` changes only
`crates/litchi-pages/src/package/table_lock.rs`. Body-table discovery performs
one coalesced proof of the rooted body’s complete field-9
(`TABLE_BODY_FIELD`) attachment inventory before walking each table’s
attachment, drawable, and model graph. That proof requires one selected
aggregate occurrence, exact field-9 declarations, and a declaration set equal
to the parsed body-table entries; the selected table still receives
independent physical-slot and message checks.

The field-9 prefix check uses checked capacity for aggregate plus field-local
object references, charges it, and `try_reserve`s one declaration map whose
counters cover both sources. The body inventory separately checked-counts all
field-9 declarations and `try_reserve`s before checking duplicates and
entry correspondence. Aggregate/field-local declaration work, declaration
and entry checks, and duplicate-prefix comparisons are charged, with each
prefix length accounted before its comparison. A wrong path, alias, duplicate
declaration, data-reference substitution, or unaccounted reference on the
selected table edge fails closed. Unrelated aggregate references on
intentionally accepted non-table paths remain valid only when their own exact
FieldInfo declaration accounts for that occurrence; this is not a
package-wide alias proof.

Strict TableInfo parent metadata remains part of the proof: the selected
TableInfo payload/metadata must carry the exact model reference and nested
body-parent path, strict selected archive/message metadata is required, and
the terminal model archive slot and supported message type are rechecked. The
changed publication carries the complete resolved `BodyTableTarget` proof
rather than only a table position. For changed rewrite and patch-apply paths,
exact source bytes and fingerprints plus the retained `BodyTableTarget` proof
and before-state are checked; the separate candidate is reopened and its
retained target after-state is verified before an immutable candidate is
returned. The inverse preserves the same proof. No-op commit/apply paths
intentionally retain unchanged-source behavior and do not claim a
changed-candidate reopen or a repeat of the full body aggregate proof.

This is Pages table-lock ownership/work hardening only. It makes no
full-transaction accounting or performance claim and admits no
Cargo/build/test or native result. It retires no `litchi-iwa` dependency edge
or ordered debt and makes no migration-host, host-exit, or monolith-exit
claim; the deletion gate is unchanged.

No native Pages open/save result is admitted by this amendment. The earlier
recorded Pages table-lock fixture was rejected by Pages as damaged, so it
remains a native-open gap rather than a pass; package-level transaction tests
and candidate reopens are not native application evidence.

## 2026-08-23 amendment: checked Keynote chart-title extension counting

Commit `0dcb85b9a` changes only
`crates/litchi-keynote/src/package/slide_chart_title.rs`. The
`read_chart_title` extension-occurrence counter now uses checked addition and
rejects counter overflow as `ChartTitleError::InvalidSource`; the existing
exactly-one extension, length-delimited wire-type, canonical-framing, and
bounded visible-title checks remain unchanged.

This is bounded source-level counting hardening for the chart-title scan only;
it does not broaden chart-title ownership or mutation semantics. No
Cargo/build/test or native result is admitted, and no performance, allocation,
dependency, debt, host, or monolith result follows.

## 2026-08-23 amendment: checked Numbers formula-category work product

Commit `28a90aa77` changes only
`crates/litchi-numbers/src/package/extractor.rs`. Its formula-category
preflight now checks the source-length product for
`MAX_FORMULA_CATEGORY_DEPTH.saturating_add(1)` passes through
`checked_formula_work_product` before applying the wire input-byte clamp;
overflow maps to the active `SemanticLimitKind::FormulaWork` ceiling rather
than saturating the product. Category traversal and retained projection
semantics are otherwise unchanged.

This is bounded source-level arithmetic hardening only. No Cargo/build/test or
native result is admitted, and no performance, allocation, dependency, debt,
host, or monolith result follows.

## 2026-08-23 amendment: Keynote transition router provenance declarations

Commit `6c7dd60dc` changes only `crates/litchi-iwa-protos/build.rs`. The
Keynote slide-transition provenance guard now requires exactly one production
codec declaration of each router field constant:
`SLIDE_TRANSITION_FIELD = 4`, `TRANSITION_ATTRIBUTES_FIELD = 2`, and
`ATTRIBUTES_ANIMATION_FIELD = 8`. These declarations are checked alongside
the canonical slide/transition/animation schema and projection-message
declarations already enforced by the guard.

This is build-time source-provenance hardening only; it does not change
runtime transition parsing, ownership, or mutation behavior. No
Cargo/build/test or native result is admitted, and no performance, allocation,
dependency, debt, host, or monolith result follows.

## 2026-08-23 amendment: Pages table-lock data-reference charges

Commit `6e1c22974` changes only
`crates/litchi-pages/src/package/table_lock.rs`. The selected-reference
helpers now charge aggregate and field-local `data_references` counts against
the bounded `PayloadReferences` budget before semantic ownership checks:
`message_declares_reference` charges the message and each field, and
`message_declares_reference_prefix` does the same for the coalesced body
proof. Selected table-edge rules and intentionally accepted unrelated
non-table paths are unchanged; this closes an accounting gap rather than
broadening ownership.

The commit adds the
`data_reference_scans_are_bounded_before_semantic_checks` low-ceiling cases
for message-level and field-level data references in both helpers, asserting
the typed `PayloadReferences` limit before semantic checks. These are
source-level budget tests; no test execution or native result is admitted,
and no dependency edge, ordered debt, host, or monolith result follows.

## 2026-08-23 amendment: neutral Numbers table-name projection

Commit `ace1d1f3a` changes only `crates/litchi-iwa/src/protobuf.rs`. The
neutral type-6001 `TableModelArchive` route now uses the strict
`numbers_names_codec::decode_table_names` projection with source-derived byte,
field, work, and nesting ceilings instead of generated
`TableModelArchive::decode`. Codec resource failures map to the existing
typed input-byte, field, rewrite-work, or nesting errors; malformed wire,
invalid UTF-8, and duplicate singular names remain invalid-format failures.

Only validated table-name text is fallibly copied with `try_reserve_exact`;
the original archive bytes remain preservation authority, and the neutral
wrapper no longer reconstructs a generated table model or publishes cell
data. Focused cases cover Unicode text, tolerated unknown fields,
malformed/truncated payloads, invalid UTF-8, duplicate names, and the absence
of a production generated decode call. This is bounded neutral-projection
hardening only: no Cargo/build/test or native result is admitted, and no
format-owner edge, ordered debt, host-exit, or monolith result follows.

## 2026-08-23 amendment: Wave34 evidence is bounded and does not advance the exit gate

The current-ref Wave34 source changes remain compatibility and bounded-work
changes, not monolith-exit work. `d24c4fb4444afa68cd739ac0e11b38ba68cde441`
hardens only the comment-storage codec fuzz target: its bounded normalizer and
ceilings are 64 KiB input, 8192 fields, 256 KiB work, 1024 references, 64 KiB
text, and recursion 64. The target cross-checks strict standalone reference
decoding against report/visitor observations, checks source borrowing and
source preservation, and includes once-per-process deterministic canonical,
malformed, callback-error, and exact Bytes/Fields/Work/References/Text/Nesting
limit probes. No fuzz campaign or sanitizer result is claimed.

`8b7a6381dfa9dd7d59812af4c24beeee65ca8070` changes only Pages footnote
projection. `native_footnotes` builds one borrowed, fallibly reserved object
location map, keeps the first object for each nonzero identifier to preserve
the prior first-match lookup, and uses it for body-footnote reference,
storage, and marker resolution. Missing identifiers still fail closed. The
commit contains focused map and patch inverse/no-op tests, but this handoff
supplies no test execution result.

`0d5a523e6d3046a2d5ab2ee18755339ff7e1501b` marks raw-ID Pages direct
drawable-comment CRUD methods and raw-ID Numbers and Keynote direct
drawable-comment/reply compatibility methods as deprecated while retaining
`IWorkDrawableCommentEditor` as the migration-host editor. Pages reply methods
are typed and carry only scoped `allow(deprecated)` internally; they are not
deprecated. The change does not alter signatures, IDs, validation,
publication, or bytes; focused text comment/reply owners remain outside the
deprecation. This is a source/rustdoc boundary, not a replacement owner or a
host-exit step.

The supplied native evidence is limited to application reopen: Pages fixture
`/private/tmp/litchi_native_pages_footnote_wave34.pages`, SHA-256
`b04a442045665fd51648342a1028976e0730f78680a3bd2b67c3f91bd40b11a`; Numbers
table-info fixture, SHA-256
`a3a34b9e374fd2cee2ac7d734f9d1693bb6ad35a3111275ae5bd0202a97e6a07`, with
formulas/comments; and Keynote transition initial/post hashes
`9c9274…`/`8e5b810…`. This is native evidence only: it establishes no Rust
parity, native save/round-trip fidelity, byte identity, or ownership move.
Native soundtrack evidence is incomplete, so no soundtrack gate or pass is
recorded.

No Cargo/build/test result is admitted beyond the supplied artifact handoff.
No `litchi-iwa` dependency edge or ordered debt is retired, and the
migration-host, host-exit, and monolith-deletion gates remain open.

## 2026-08-24 amendment: native Keynote chart-title application gate

The supplied Keynote application gate used fixture
`/private/tmp/litchi_native_keynote_chart_title_gate_wave35.key`, final
SHA-256
`f50a8efb2bef885b3759b4bcd5e78f505da0074492774a8dbdde4e8ea9b07e56`. In
Keynote, a real 2D Column chart's title was set to `Native Chart Title Gate`;
after save, close, and reopen, the title persisted without a warning or repair
prompt. The title was then cleared; after save, close, and reopen, it was
absent, again without a warning or repair prompt. The source fixture
`test-data/iwork/keynote/basic.key` was unchanged at SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Keynote was quit after the gate.

This is application-only evidence. It establishes no Rust/native parity, no
ownership result, no migration-debt retirement, and no `litchi-iwa`
monolith-exit or deletion claim.


## 2026-08-23 amendment: current-HEAD scope correction

Review of committed HEAD `baa1cadd` confirms that the dated evidence remains
narrow. The `28a90aa77` formula-category change checks only the source-length
product for `MAX_FORMULA_CATEGORY_DEPTH.saturating_add(1)` passes through
`checked_formula_work_product` before the wire input-byte clamp. It does not
claim that category traversal, retained projection, or every formula work
product was newly proven; the separate `711e2b517` wording covers only the
formula-reference, TableInfo, and table-model-name products and explicitly
makes no category-scan claim.

The Pages `FootnoteSemanticBudget` is per
`project_body_footnotes` projection or `native_footnotes` pass. It charges
projected footnote text and custom-marker bytes before per-value ownership, but
does not charge staged `set` text, `after` values, or candidate-package
retention. `TransactionBudget` belongs to the section transaction/rewrite
paths; the footnote-text rewrite seam has separate wire/text limits. Neither
budget wording is a package-wide or full-transaction accounting claim.

Numbers message-type wording is context-scoped. In strict rooted ownership,
type `6000` is the TableInfo owner and type `6001` is the TableModel
payload; the global compatibility projection may also inspect a
model-shaped legacy type-`6000` payload only after its shape gate. The neutral
IWA registry's `6000`/`6001` table-model entries are a bounded name-only
projection and do not publish cell data or establish complete model migration.
The comment ownership census can validate referenced type-`6011` segment
messages, but the recorded native fixture contained zero type-`6011` entries;
there is therefore no native Segment-support or type-`6011` proof.

The `baa1cadd` source change adds bounded comment-cell scan work and shared
comment-text ownership only. It supplies no new Cargo, native, Segment,
dependency-edge, host-exit, or monolith-deletion evidence. This amendment is
documentation-only and does not alter any historical tail.

## 2026-08-23 amendment: current HEAD bounded follow-up

The prior dated evidence and scope notes remain historical as written. At
committed HEAD `4f2c14484` (parent `609d44049`), the post-`b208dbd77` changes
are bounded source, test, deprecation, and fuzz-harness hardening:

- `d131176b4`, `94b07ffc5`, `03cee9a79`, and `7a497d1ef` retain and ratchet the archive
  production Buffa lazy-view boundary by rejecting generated eager helpers;
- `6b705d39b`, `f8060a43f`, and `447cb3ec3` keep duplicate unselected/shared
  Keynote slide owners rejected before read or staging and add focused
  coverage;
- `33656a7f0` and `8254df4e9` preserve typed Pages footnote wire and
  semantic/rewrite limit categories;
- `cd2dc0ca5` deprecates only the legacy raw-ID Pages section-clear API;
- `d3b1bd571` and `5abf10bb4` add mixed reorder/replacement cache-publication
  coverage;
- `95356f2a8` rejects case-insensitive worksheet-name duplicates;
- `543f74a37` makes Numbers table-header growth fallible per iterator item;
- `cca10f240` redacts native reference identifiers from diagnostic formatting;
- `7effedb57` adds a deprecation boundary to legacy raw-ID Numbers
  cell-comment APIs while retaining compatibility and migration-host paths;
- `609d44049` fixes BIFF8 supplementary-character length accounting and adds
  validation-string round-trip coverage;
- `4f2c14484` charges nested table-lock references before declaration-map
  reservation and adds a low-ceiling budget case;
- `b5a9d5994` bounds CFB fuzz input to 4 MiB, while `1446dc8e7` bounds RTF
  fuzz inputs to 1 MiB and parser/opaque work; both supply campaign recipes
  only.

No test/build/Cargo execution, sanitizer campaign, native application
result, performance/allocation measurement, dependency-edge or ordered-debt
change, host-exit, or monolith-deletion result is admitted by this amendment.

## 2026-08-23 amendment: current HEAD truth audit

The preceding `4f2c14484` note is a historical snapshot, not the current
monolith status. At committed HEAD `51ca8eae7` (parent `9366d0a00`),
`b316320c9` deprecates the `litchi-iwa::raw` compatibility facade and updates
ordered debt 9's exit wording to include removal of that facade. The host,
its dependency edge, and all 13 ordered debts remain present. The Keynote
owned-title setter change (`e4533c17d`), XLS test-call cleanup (`9366d0a00`),
and fallible Pages footnote staging copies (`51ca8eae7`) do not move ownership
out of the migration host or satisfy a deletion-gate item.

The archived HEAD still inventories 64 workspace packages and 239 internal
dependency declarations. Its complete boundary checker remains non-green with
22 findings: two Keynote slide-transition validator exports, 18 Pages
table-lock alias sites, and the two existing `TableLockState`/`TableSelector`
aliases. These are current source-audit results, not an exit or deletion
gate. No test/build/Cargo execution, native application or Rust/native parity
result, Buffa/Prost migration, generated-schema retirement, dependency-edge or
ordered-debt retirement, host exit, or monolith deletion is admitted.

## 2026-08-23 amendment: current HEAD extension after the truth audit

The preceding `51ca8eae7` audit is historical. At current HEAD
`ff782f2f5` (parent `dab07a09f`), the Keynote chart-title owner guard now
rejects shared title stand-ins before read or edit (`dab07a09f`), while the
migration-host `KeynoteSlideInfo` title/body/speaker-notes native IDs are
explicitly deprecated (`ff782f2f5`). Both changes retain compatibility code;
neither moves an owner out of `litchi-iwa` or retires a deletion-gate item.

The host, its dependency edge, and all 13 ordered debts remain. No test/build
execution, native application or Rust/native parity result, Buffa/Prost
migration, generated-schema retirement, dependency-edge/debt retirement,
host exit, or monolith deletion is admitted. The current boundary audit
remains non-green with 22 findings: two Keynote slide-transition validator
exports, 18 Pages table-lock alias sites, and the two existing
`TableLockState`/`TableSelector` aliases.

## 2026-08-23 amendment: current HEAD extension after cache and stale-metadata fixes

The preceding `ff782f2f5` note is historical. At current HEAD `8cddd44af`,
`09286432e` tightens the Pages table-lock boundary by charging empty body
metadata before declaration reservation and rejecting stale table ownership;
`8cddd44af` adds only a moved/replaced-entry cache test. Neither change moves
ownership out of `litchi-iwa` or satisfies a deletion-gate item.

The host, its dependency edge, and all 13 ordered debts remain. No test/build
execution, native or Rust/native parity result, Buffa/Prost migration,
generated-schema retirement, dependency-edge/debt retirement, host exit, or
monolith deletion is admitted. The boundary checker remains non-green with
22 findings: two Keynote slide-transition validator exports, 18 Pages
table-lock alias sites, and the two existing `TableLockState`/`TableSelector`
aliases.

## 2026-08-23 amendment: current HEAD extension after deprecation and alias guards

The preceding `8cddd44af` note is historical. At the then-current committed
HEAD `546df8f04`, `406cff111` adds boundary ratchets for retained Numbers and
Pages raw-ID compatibility methods, and `546df8f04` rejects aliased Keynote
slide/show/node identities before transition reads or staging. These changes
retain the migration-host compatibility surfaces and move no owner out of
`litchi-iwa`; no deletion-gate item is satisfied.

The host, its dependency edge, and all 13 ordered debts remain. No test/build
execution, native or Rust/native parity result, Buffa/Prost migration,
generated-schema retirement, dependency-edge/debt retirement, host exit, or
monolith deletion is admitted. The boundary checker remains non-green with
22 findings: two Keynote slide-transition validator exports, 18 Pages
table-lock alias sites, and the two existing `TableLockState`/`TableSelector`
aliases.

## 2026-08-23 amendment: Wave45 source-only follow-up at HEAD 97972ce74

The preceding `546df8f04` note is historical. At committed HEAD
`97972ce74` (parent `875fe2a48`), the Wave45 changes do not advance the
monolith deletion gate:

- `6d7f1a701` repairs the Pages native-message provenance guard markers. The
  supplied direct Pages guard evidence is `16/16`; it is scoped guard
  evidence, not a full-workspace gate.
- `1bad9156e` tightens the strict Numbers TableDataList/Segment route probe so
  known repeated entry/segment fields must carry length-delimited wire shape
  before a candidate is admitted to full decode. This is source-level wire
  hardening; no test execution is recorded here.
- `875fe2a48` stages private bundle aggregate fields/work/text/output budgets
  with transactional publication. The seam is explicitly `dead_code`-allowed
  pending neutral-decoder cutover and does not widen the 6005/6201/6011
  compatibility routes. Review was limited to rustfmt/diff; Cargo was blocked
  before this repair, so no Cargo or test result is claimed.
- `97972ce74` repairs the Keynote soundtrack projection digest and adds a
  scalar provenance guard. The supplied soundtrack-focused `cargo check`
  passes; the full `litchi-iwa-protos` test remains blocked by the
  `table_info` compile failure.

No native or type-6011 result, Rust/native parity, full-workspace verification,
Buffa/Prost exit, generated-schema retirement, dependency/debt retirement,
migration-host exit, or monolith deletion is admitted by this amendment.

## 2026-08-24 amendment: accepted bounded follow-up at `2bbf3c64`

The historical monolith/deletion tails above are retained. The accepted chain
does not move an owner out of `litchi-iwa` or satisfy a deletion-gate item:

- `7468fdfc4` preserves copy-on-write identity for an exact-byte
  `EntryStore::replace_data` no-op, with a focused `Arc::ptr_eq` case.
- `322032c10` deprecates three raw-ID Keynote chart-title methods and adds
  typed-selector boundary ratchets; the migration-host compatibility methods
  remain.
- `031678a31` adds strict alternate type-`6201` TableDataList coverage and
  rejects a non-canonical duplicate scalar. Its type-`6011` data is synthetic
  test input, not native Segment evidence.
- `2bbf3c64` shares the Pages footnote source/candidate semantic budget and
  rejects observed `19` at exact limit `14`; this is local projection
  hardening, not a workspace memory or performance result.

The current lightweight checks are `git diff --check` and
`python3 -m unittest tools.test_check_crate_boundaries` (314 tests). No
full-workspace Cargo gate is claimed because the surrounding checkout has a
dirty/staged snapshot and `Cargo.lock` is not tracked by this parent. Clean
archived Buffa/protos results remain historical only: 33 focused protocol
tests and a full `litchi-iwa-protos` baseline at 267 passed/0 failed, with a
later strict Numbers/Buffa report at 237 passing protos tests. No accepted
commit removes Prost or changes generated-schema ownership.

Prior reports supply only historical native/fuzz context: a Numbers
comment/formula/rich-text reopen recorded 24 type-6005 and zero type-6011
entries, and TableDataList/CommentStorage smokes recorded 100 ASan runs with
11/20 seeds. No native type-6011 proof or current fuzz campaign follows.
The host edge, all ordered debt, and the monolith remain; no dependency-edge
or debt retirement, migration-host exit, or deletion claim is made.

## 2026-08-23 amendment: Wave51 Keynote boundary ownership correction

Current opaque Keynote transition wire validation belongs to the doc-hidden
`litchi-iwa-protos::keynote_slide_transition_codec`, not the focused public
transition facade. `litchi-keynote` continues to own selector-first package
transactions, rooted graph admission, exact raw-byte patching, inverse proof,
and error mapping. `litchi-iwa` remains a compatibility consumer and retains
its semantic conversion adapter. The removed
`litchi_keynote::transition::validate_opaque_transition_settings` function was
adapter-only and intentionally source-breaking for the unpublished `0.0.1`
surface; no public wire-oriented compatibility shim is retained.

The same slice makes focused `litchi-keynote::Package` exact-file opening
descriptor-first, no-follow/reparse/device-rejecting, mutation-checked, and bounded.
That admission hardening does not move a package owner out of the migration
host, replace archive-free `Document` preparation, or establish atomic/durable
filesystem publication.

The two Keynote boundary findings are gone. A clean tracked audit still has 20
unrelated focused Pages table-lock aliases; the shared worktree adds three
findings from its untracked Pages host table-lock source. Debt 010
(`litchi-iwa -> litchi-iwa-protos`) and debt 014
(`litchi-iwa -> litchi-keynote`) both remain, as do all other ordered debts.
The inventory is unchanged at 64 packages, 239 internal declarations, one
migration host, and 13 debts. This is boundary and ingress hardening, not
generated-schema or Prost retirement, dependency-edge removal, native
application evidence, host exit, or monolith deletion.

## 2026-08-23 amendment: Wave52 Numbers formula projection (not a monolith-exit gate)

The preceding monolith/deletion evidence remains historical. The current
Numbers formula slice changes only the production formula reader boundary: the
bounded owned `FormulaArchiveBytes` source copy remains the preservation
authority, strict schema preflight charges the admitted source before a private
Buffa lazy per-node/event stream, and that stream replaces the per-cell
generated `FormulaArchive` graph. No generated repeated/LazyRepeatedView is
materialized and no `to_owned_message` conversion is used in the production
formula path. Within the `litchi-numbers` formula path, generated Prost
messages remain a test-only oracle for fixture construction and differential
checks.

This does not establish workspace-wide Prost freedom. Commit `2c538d496`
moves the direct `litchi-numbers` `prost` declaration from normal dependencies
to dev-dependencies; its generated `FormulaArchive` builders, decoder, and
reference renderer are `cfg(test)`-gated oracles. That removes the focused
normal dependency edge without removing Prost from other workspace owners. It
does not move the migration host owner, remove the `litchi-iwa` dependency
edge, retire an ordered debt, or satisfy a host/monolith deletion gate. It also
makes no mutation, publication, or durable-output change.

The exact isolated verification used a detached snapshot plus a temporary,
uncommitted bypass of the unrelated Pages provenance guard:

- `cargo test -p litchi-iwa-protos numbers_formula_codec --lib`: 36 passed;
- `cargo test -p litchi-numbers --lib package::extractor::tests::compatibility_`:
  7 passed;
- the corresponding `raw_formula_archive_` and `scalar_formula_` filters:
  9 passed each;
- `cargo check -p litchi-iwa-protos -p litchi-numbers --all-targets`: passed;
- `cargo clippy -p litchi-iwa-protos --lib -- -D warnings`: passed. Strict
  Numbers Clippy reached only two untouched deprecated `object_count` uses
  and one untouched `manual_contains` finding;
- `python3 -m unittest tools.test_check_crate_boundaries`: 319 passed, and the
  tracked audit retained 20 Pages findings and zero Numbers findings.

Broader formula tests were also attempted, but existing table-storage and
table-transaction verification failures outside the migrated reader prevent a
full-suite green claim. On macOS, Numbers 14.4 opened the read-only native
fixture `/private/tmp/litchi-wave52-formula-native.2dgQQv/source.numbers`
without repair. Its visible 1,200 by 8 table exposed 2,398 formulas, result `5`
at E2, and the intentional division-by-zero at F2. Formula-editor inspection
was cancelled. The Rust `read_numbers` example reported one rooted sheet, one
1,200 by 8 table with 9,600 materialized cells, and one compatibility table.
SHA-256 remained
`e0fb395b0e819583f14d72f1d4916ce482ff35dce9de895d3dfd26869888bbc9`
through both read-only checks.

No exact RSS/performance, native save/reopen mutation, Rust/native formula
parity, workspace-wide Prost removal, host exit, debt retirement, or monolith
deletion follows from this bounded source migration.

## 2026-08-24 amendment: Wave53 Keynote slide-background bounded owner slice (not a monolith-exit gate)

Commit `0b12df5e1` moves the selector-first Keynote slide-background semantic
transaction into `litchi-keynote` while retaining `litchi-iwa` as a
compatibility bridge. The package owns effective/direct background reads,
solid/gradient/none/opaque setters, inheritance reset, exact reversible
patches, copy-on-write style handling, stylesheet culling, package-metadata
UUID/external-reference updates, and atomic preview-aware publication. The
doc-hidden `litchi-iwa-protos::keynote_slide_background_codec` owns the strict
borrowed wire snapshots and bounded typed rewrites. The host retains its
compatibility surface and does not become the format owner again.

The focused verification is 25/25 Keynote package cases, 19/19 background
codec cases, 13/13 package-metadata codec cases, and 9/9 migration-host
background cases. Strict Keynote/protos Clippy passes, and the all-target
check for `litchi-iwa-protos`, `litchi-keynote`, and `litchi-iwa` passes with
the checkout's existing warnings. Boundary units pass 324/324; the live
explain audit exits 1 only for 23 known Pages findings (20 tracked and three
untracked), with no Keynote finding. Targeted formatting and diff checks pass.

The manifest and workspace ownership claims are unchanged: the focused
Keynote direct Prost edge is dev-only, `litchi-iwa-protos` retains normal Prost
for generated owners, and the Numbers formula oracle remains dev-only. This
slice retires no `litchi-iwa` dependency edge or ordered debt, and it does not
retire generated schemas, remove the migration host, or delete the monolith.

The bounded native acceptance used Keynote 14.4 on disposable candidates from
the pristine text-only `test-data/iwork/keynote/basic.key` fixture (500,058
bytes; SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`). Keynote
opened the candidates without repair. The pre-save candidate hashes/sizes
were solid `da9437a0ee66a71cea632c427c41fdb45ff051e8e7fcf274cd841ed46bd76b82`
(456,230 bytes), gradient
`425a189d95cb535b5676d999f3e3d811c710de925c524a8e9877968e9a9394d5`
(456,251), none
`f50470a2356c188a14d27c4ab995e19c609d052fbf0716af4b7876da72feff83`
(456,209), and reset
`e171028d914283d90146cc7fc3ade20f54262c21d189cc6a638cfee9a9d91852`
(456,207). Keynote displayed dark red Color Fill, red-to-blue Gradient Fill
at 45 degrees, No Fill, and reset White at 100 percent; visible text stayed
unchanged. After save/close/reopen, normalized outputs were solid
`c2748466dda358870df9be1c8ba5c8e390a1cef86a136d73bee1d788efa1ec9b`
(499,471 bytes), gradient
`d41ea6c6e5aa3224824bac8461460c625cfbbc6870e9c8320b45b0ef705a4cd7`
(503,600), none
`15f0d2da655c573e80661893aac3fb0f4468704596509699c671545abafa4f62`
(468,839), and reset
`ae299a456bf40af628c6f8dc900cdc1e90c3cf1848de6a534a214874e813056a`
(500,039).

This is bounded application acceptance on a text-only fixture only. It does
not establish Rust/native parity, media preservation, exact RSS or allocation
behavior, performance, or a durable workspace publication gate.

## 2026-08-24 amendment: Wave54 Numbers table-cell storage bounded reader slice (not a monolith-exit gate)

Commit `b7f720872` moves strict `TableDataList` and
`TableDataListSegment` read projection into the doc-hidden
`litchi-iwa-protos::numbers_table_cell_storage_codec`, but the three migrated
helpers and the package graph still live in the `litchi-iwa` Numbers editor.
The migration host continues to own candidate routing, aggregate operation
budgets, segment lookup, semantic result staging, and all generated mutation
and writer paths. No public API or manifest edge changes.

The focused codec/host regressions, all-target check, scoped strict Clippy, and
335/335 boundary-policy tests pass. The live boundary command remains
non-green only for 23 known Pages findings and reports no Numbers storage
finding. Numbers 14.4 also opened, saved, closed, and reopened the disposable
sorted-text candidate without repair while preserving the visible stable
equal-key order; ADR 0008 records the exact hashes and bounded caveats.

Generated `TableDataList` schemas, Prost-backed mutation code, the
`litchi-iwa` owner and dependency edges, all ordered debts, and the migration
host remain. This slice retires no generated schema or Prost use, moves no
package owner out of the host, and satisfies no host- or monolith-deletion
gate.

## 2026-08-24 amendment: Wave55 PackageMetadata registry reader slice (not a monolith-exit gate)

Commit `02eeb3acc29801838e3981236fd1e41ecbc44fee` cuts only three private
PackageMetadata registry readers from owned generated-message projection to the
doc-hidden `litchi_iwa_protos::package_metadata_codec` borrowed visitor:
`component_identifier_for_entry` has 157 non-definition callers,
`component_identifier_for_object_uuid` 14, and `component_uuid_identifiers` 41
(209 total). `litchi-iwa` retains the compatibility query surface and package
ownership; mutation, writer, allocator, save-token, and data-reference routes
remain generated Prost routes.

The visitor's two-pass finite resource admission, strict selected-projection
canonicality, fallible UUID-set staging, and typed limit/allocation failures
are recorded in ADR 0005. Focused codec, host, caller, check, Clippy, boundary,
format, and documentation gates are recorded in ADR 0008. The live boundary
baseline remains only the known 23 Pages table-lock findings. Numbers 14.4
accepted the duplicated-sheet candidate through save/close/reopen in the UI,
but native normalization leaves a known duplicate-component-identity
`InvalidFormat` reread debt, so this is application acceptance only and not
Rust/native post-save parity.

No manifest or public-API change, package-owner move, generated-schema or
Prost retirement, ordered-debt retirement, migration-host exit, or
`litchi-iwa` monolith-deletion claim follows from this bounded reader slice.

## 2026-08-24 amendment: Wave56 Numbers exact-alias bounded owner slice (not a monolith-exit gate)

Commit `e4952f7eeaa196263a28fac6cf86510c3e11a2f6` changes only the
focused `litchi-numbers` physical ingress/index boundary, plus the neutral
`litchi-iwa-core` source-content comparison it relies on. Exact copies of one
identifier in distinct components may share one logical bare-ID resolver
entry; the physical components and their bytes remain intact. Same-component,
divergent, raw-header-divergent, and framing-divergent duplicates remain
errors. The public focused API gains no raw object identity, and the boundary
ratchet scans the complete Numbers package owner for any such reintroduction.

This closes the Wave55 Rust reread debt for the observed eight exact aliases,
not general component-qualified identity. The current resolver deliberately
cannot choose between divergent owners. Mutating an aliased object therefore
fails candidate reopen atomically, while unaliased semantic edits preserve the
physical aliases. Owner-aware component routing, PackageMetadata provenance
and save-token publication, and durable native resave remain future work.

The focused core/package, all-target, scoped Clippy, and 361/361 boundary gates
are recorded in ADR 0008. Numbers 14.4 opened and rendered the Rust and
Numbers-emitted candidates, and the strict reader reopens both, but repeated
autosave/resave returned `TSPersistence` code 2 while a pristine same-directory
control saved. This is bounded reader/application evidence only and is not a
native publication gate.

No public semantic facade or manifest edge changed. `litchi-iwa` host and
dependency edges, generated schemas, normal Prost owners elsewhere, ordered
debts, migration-host responsibilities, and the monolith all remain. This
slice retires none of those owners or edges and satisfies no host-exit or
monolith-deletion gate.

## 2026-08-24 amendment: Wave57 Numbers name-publication bounded owner slice (not a monolith-exit gate)

Commit `35c2ae281d40bc16c659c30dd3690a702a750ea0` adds one bounded semantic
publication route to Numbers name edits. Follow-up accounting hardening is
commit `ba57356166e82ff37ee1cd5223b68acd484aac42`; it removes detached,
unmetered candidate-size `Budget` work and makes candidate sizing and selector
comparisons share the reported `Budget`. `litchi-numbers` owns the transaction
and its exact source/target patch artifacts; the hidden
`litchi-iwa-protos::package_metadata_codec` owns strict borrowed save-token
projection and validation. A changed transaction requires one exact
type-11006 `Index/Metadata.iwa`, advances the root token once, and updates only
selected current components matched by identifier plus effective locator.
Versioned and unselected records, root `last_object_identifier`, and unknown
metadata bytes remain unchanged. No-op edits bypass metadata and remain
byte-exact; inverse application restores the original package bytes. The
metadata member is an additional sidecar rewrite and is excluded from the
semantic native `touched_components` count.

The operation's source/physical-entry, metadata, selector, field, depth, work,
component, output-size, and candidate-verification budgets are local to this
route. Exact sizing and fallible reservation precede the single candidate
allocation; a test-only aggregate proves successful budget charges equal
`RewriteReport.work_bytes`, and max-minus-one fails before allocation. Before
native rewrite/allocation, the Names caller precharges decompressed type-11006
payload bytes times the changed semantic-operation upper bound for visitor
locator matching; compressed Snappy bytes are separate publication accounting.
This does not add a public raw-ID API, remove a manifest edge, or
retire generated schemas or normal Prost owners; it is not evidence for a
workspace-wide dependency or performance claim.

Focused verification is recorded in ADR 0008: codec 19/19, names 19/19,
Numbers package 37/37, boundary tests 369/369, and the scoped checks passed
with the documented unrelated baselines. The live boundary checker retains
exactly the 23 known Pages findings, and the full Numbers library retains one
unrelated storage canonical-framing failure. No full-workspace green claim is
made.

Numbers 14.4 opened and rendered the final Rust candidate from
`/private/tmp/litchi-wave57-final-native.x3zj8I/rust-renamed.numbers` (82,283
bytes, SHA-256
`770a4d593ceede1b14fdee7803d6de780c33d435f270ef572341e6cf4ec37830`) without
repair; its inverse exactly matched the normalized duplicate source
`/private/tmp/litchi-wave55-metadata-native.4D1yTa/duplicated.numbers`
(139,219 bytes, SHA-256
`790aa7386ad6f5bb641dda0aaef3a47236cbe727b9712a5a9cf581ffce6d6754`). A
Numbers Save As emitted an artifact despite the known TSPersistence code-2
alert, and that artifact reopened/rendered without repair. The normalized
source also exhibits the alert while a pristine no-alias control saves; this
is bounded open/render and emitted-artifact reopen/readback evidence only, not
native save acceptance, durable publication, Rust/native parity, or native
token-set acceptance.

This amendment retires no `litchi-iwa` host edge, ordered debt, compatibility
responsibility, generated schema, or monolith owner. The migration host and
its existing host/edge/debt obligations remain, and this slice satisfies no
host-exit or monolith-deletion gate.

## 2026-08-24 amendment: Wave58 focused Numbers comment replacement (not a monolith-exit gate)

Implementation commit `e96a3a301` moves the supported existing-root
cell-comment replacement publication through the selector-first
`litchi-numbers` package owner. Native table identifiers remain confined to
the deprecated migration-host boundary and are converted to scoped sheet and
table selectors before the focused call. The focused crate receives no raw
object identifier and adds no public raw-ID method or typed native-ID
parameter.

This is an intentionally bounded delegation, not a host deletion. Comment
creation, replies, shared ownership, segmented storage, clear/graph cleanup,
and sources whose broader table projection is not yet accepted remain in the
compatibility host. The migration host still owns its public deprecated API,
native selector adapter, fallback writer, and exact post-publication reopen.
No manifest edge, generated schema, Buffa/Prost owner, ordered debt, or other
`litchi-iwa` responsibility is retired. Wave58 therefore satisfies no
monolith-exit, host-removal, dependency-removal, or publication-completeness
gate.

## 2026-08-24 amendment: Wave60 Numbers root comment-clear bounded owner slice (not a monolith-exit gate)

The Wave60 implementation series culminates in commit
`895ef17848516cf201e717239c44c119eae5da87`. `litchi-numbers` now owns the
selector-first clear of one exact, unshared, reply-free root cell-comment
graph, including cell/list mutation, unshared storage deletion, archive
reference pruning, Metadata save-token publication, preview invalidation,
candidate verification, and exact patch/inverse artifacts. Supporting neutral
archive-reference and doc-hidden PackageMetadata codecs provide strict,
source-preserving inspection and rewrite seams without exposing raw object IDs
through the semantic API.

This is intentionally narrower than general comment deletion. Shared or
segmented storage, replies, exact aliases, unknown archive owners, registered
Metadata UUID/external/data/root-map ownership, missing or ambiguous Metadata,
comment creation without an existing strict comment list/author graph, and
wider graph cleanup remain refused by the focused owner.
The deprecated `litchi-iwa` raw-ID clear still exists: metadata-bearing
accepted sources delegate through semantic selectors, while sources with no
Metadata sidecar retain the legacy compatibility writer. Focused errors on a
metadata-bearing source fail hard rather than falling back.

The focused codec/package/host and 387/387 boundary gates, along with bounded
Numbers 14.4 open/render/save/close/reopen evidence, are recorded in ADR 0008.
The live checker retains exactly the 23 known Pages table-lock findings. No
full-workspace green, arbitrary comment-graph publication, native byte-parity,
or performance claim is made.

No manifest edge, generated-schema owner, normal Prost owner, ordered debt,
migration-host API, or `litchi-iwa` crate responsibility is retired. The host
adapter, no-Metadata fallback, unsupported comment graphs, Pages table-lock
baseline, and the monolith itself remain. Wave60 therefore satisfies no
host-exit, dependency-removal, or monolith-deletion gate.

## 2026-08-24 amendment: Wave61 Pages body-table lock bounded owner slice (not a monolith-exit gate)

Commits `a27a22c9cf95d548da7634167d6bac7190b8321a` through
`ca3fbd21a24f7195ef9b2d8d169e286339fc274e` harden the selector-first
`litchi-pages` body-table lock owner, accept the aggregate-only reference
metadata emitted by Pages while rejecting contradictory ownership, remove the
unpublished flat semantic aliases, and improve operation-local work preflight.
Commit `4742e20107f29a1990d6a1886d8046a9333133b5` prevents alternate aliases,
wildcard exports, renamed host methods, and relocated host modules/examples
from bypassing the boundary ratchet.

The focused public vocabulary now contains only `BodyTableSelector`,
`BodyTableLockState`, `Package::{body_table_lock, edit_body_table_lock,
apply_body_table_lock}`, and the `BodyTableLock*` transaction types. Raw
object IDs, archive routes, wire values, and generated schema types remain
private. Pages 14.4 opened the Rust-locked candidate without repair,
recognized it as locked, saved and normalized it without an alert, and
preserved the lock across close/reopen; the strict Rust owner reread that
artifact. This is bounded acceptance for this operation, not byte parity or a
general Pages publication gate.

The `litchi-iwa` Pages compatibility host, its dependency edges, ordered
debts, and the monolith remain. The live boundary checker still reports the
three user-owned untracked retired-host findings, so this amendment does not
claim host retirement. No manifest edge, generated schema, Buffa/Prost owner,
or other migration-host responsibility was removed. Wave61 therefore
satisfies no host-exit, dependency-removal, or monolith-deletion gate and
makes no package-wide performance or full-workspace green claim.

## 2026-08-24 amendment: Wave62 Pages body-table title bounded owner slice (not a monolith-exit gate)

Commits `48f203aae56e43133fd931accfa6558661594ba0` and
`a1c1e83a3edad808bacefa648fe0c1bdd53308f5` establish and harden the
selector-first `litchi-pages` body-table title owner. Commit
`a92f8f11a50c709877b8d1f0a158da72114dc4ab` removes the tracked raw-`u64`
Pages title host methods and migrates the retained examples, while commit
`a7088be4dd9fde9b2e843473b093839e7b7629b3` ratchets the complete focused
source tree, host, examples, and README against their return. The public
focused seam uses `BodyTableSelector` and presence-preserving semantic title
settings; raw IDs, wire values, archive routes, and generated types remain
private.

Pages 14.4 opened the Rust title/outline candidate without repair and
preserved the visible Title/Outline/Caption state across close/reopen. The
Rust transaction changed one native component, removed the three root
previews, and produced an exact inverse; Pages subsequently normalized the
candidate bytes. This is bounded acceptance for this semantic operation, not
byte parity or a general Pages publication gate.

This slice retires one tracked Pages compatibility-host operation, but the
`litchi-iwa` crate, its remaining Pages and Keynote hosts, dependency edges,
ordered debts, generated schemas, Buffa/Prost owners, and the monolith remain.
The live boundary checker also retains exactly three findings from the
user-owned untracked Pages table-lock host file. No manifest edge or other
migration-host responsibility was removed. Wave62 therefore satisfies no
crate-exit, workspace dependency-removal, generated-schema retirement, or
monolith-deletion gate and makes no package-wide performance or full-workspace
green claim.

## 2026-08-24 amendment: Wave63 Pages body-table header bounded owner slice (not a monolith-exit gate)

The Wave63 series from `6f35a3d88` through final hardening commit
`7bc68903ef8f107b4473843945c908856988215f` establishes the selector-first
Pages body-table header/footer owner, the doc-hidden neutral strict codec, the
focused package transaction, the tracked host cutover, and the boundary
ratchet. The public seam uses `BodyTableSelector`, semantic
`table::headers::{Count, Settings}`, and
`Package::{body_table_header_settings, edit_body_table_header_settings,
apply_body_table_header_settings}`. Raw IDs, archive routes, wire values, and
generated types remain private.

The transaction preserves optional-field presence and unknown source bytes,
rejects malformed or external references and unsupported dependency graphs,
invalidates the root previews on changed publication, reopens the complete
candidate, and retains exact source/target inverse artifacts. Pages 14.4
opened the Rust candidate without repair, rendered the 5-by-4 table with one
header row, saved, closed, and reopened it; strict Rust reread preserved the
requested frozen-row semantic after Pages normalized the package. The Pages
UI did not directly expose that freeze value, so this is bounded
open/render/save/reopen plus Rust-reread evidence, not a visual freeze-state
claim or byte parity.

The tracked Pages raw-ID header methods are retired, but the `litchi-iwa`
crate, remaining Pages dimensions/appearance/hidden-axis and footnote work,
other migration hosts, dependency edges, ordered debts, and the monolith
remain. No manifest edge, generated schema, Buffa/Prost owner, or other
workspace dependency was removed. The live checker also retains the three
user-owned untracked Pages table-lock findings. Wave63 therefore satisfies no
crate-exit, dependency-removal, generated-schema-retirement, or
monolith-deletion gate and makes no performance, general publication, or
full-workspace green claim.

## 2026-08-24 amendment: Wave64 Keynote chart-caption bounded owner slice (not a monolith-exit gate)

Commit `514b82bdf658d78ea0154f4fc075b1b85488f31c` adds the
selector-first `litchi-keynote` chart-caption replacement transaction, reuses
the strict internal caption/text codecs, delegates existing-caption host reads
and replacements through semantic selectors, and ratchets the focused facade
against raw-ID or legacy-call regressions. Exact patches, candidate reopen,
unknown-field preservation, exclusive storage ownership, and native Keynote
14.4 save/reopen acceptance are recorded in ADR 0008.

This slice deliberately does not own caption graph creation or removal.
`litchi-iwa` still allocates native caption objects, registers their graph, and
applies native stand-in removal policy for compatibility callers. The
deprecated raw-ID methods, remaining Keynote chart/style/text hosts, archive
codecs, generated schemas, normal Prost owners, dependency edges, ordered
debts, and the `litchi-iwa` monolith all remain. The live checker also retains
the three user-owned untracked Pages table-lock findings.

No manifest edge, generated-schema owner, migration-host crate, or workspace
dependency was removed. Wave64 therefore satisfies no crate-exit,
dependency-removal, generated-schema-retirement, or monolith-deletion gate and
makes no arbitrary-caption-graph, performance, or full-workspace green claim.

## 2026-08-24 amendment: Wave65 Keynote chart-caption full owner slice (not a monolith-exit gate)

Commit `f0bbe079b094b3652751f6ab7c89b2dc64fac6a9` moves the admitted
canonical Keynote chart-caption graph lifecycle into `litchi-keynote`:
stand-in-to-inline creation, existing-text replacement, active-to-fresh-
stand-in removal, Metadata UUID/save-token registration, preview invalidation,
candidate reopen, and exact patch/inverse artifacts. The production
`litchi-iwa` chart-caption host now exposes only selector-based methods and
the tracked raw-ID getter/setter/remover and graph helpers were retired; the
retained example uses the semantic selector owner.

The focused package still fails closed for cross-component, shared,
ambiguous, malformed, hostile-diff, or unsupported future graph shapes. This
is full ownership of the admitted canonical lifecycle, not a claim that every
Keynote caption graph is publishable. The native record is bounded Keynote
14.4 open/render/save/close/reopen evidence; focused Rust readback of the
normalized native candidates returns `InvalidSource`, so it is not a Rust
re-ingress, byte-parity, arbitrary-graph, or publication-completeness gate.

The `litchi-iwa` crate, remaining Keynote/Pages/Numbers compatibility hosts,
dependency edges, ordered debts, generated schemas, Buffa/Prost owners, and
monolith remain. No manifest, crate, host, generated-schema, or normal Prost
exit was claimed. The live boundary checker retains exactly the three
user-owned untracked Pages table-lock findings and no Keynote caption finding.
Wave65 therefore satisfies no crate-exit, host-exit, dependency-removal,
generated-schema-retirement, monolith-deletion, full-workspace-green, or
package-wide-performance gate.

## 2026-08-24 amendment: Wave66 Keynote chart-caption hardening (not a monolith-exit gate)

Commit `cd76c394e2cd6e360cd01e8bd61cb7594d837701` hardens the focused
Keynote chart-caption owner with source-authoritative chart-reference
transitions, canonical IWA object framing, strict selected wire validation,
exact current Metadata save-token updates for existing text, and proven
cross-component stylesheet/paragraph external references. The semantic public
API and the Wave65 host cutover remain unchanged; no raw-ID compatibility
surface was restored.

This is a bounded acceptance and safety refinement. Arbitrary future caption
graphs, ambiguous/weak/versioned dependencies, unrelated owners, and
noncanonical selected wire forms remain rejected. The native record is
bounded Keynote 14.4 open/render/save/close/reopen evidence with an exact
pre-native inverse; it is not byte parity, arbitrary graph acceptance,
package-wide performance, or a general publication gate.

The `litchi-iwa` crate, remaining Keynote/Pages/Numbers compatibility hosts,
dependency edges, ordered debts, generated schemas, Buffa/Prost owners, and
monolith remain. Wave66 removes no manifest edge, crate, remaining host,
generated-schema owner, or normal Prost owner. It therefore satisfies no
crate-exit, host-exit, dependency-removal, generated-schema-retirement,
monolith-deletion, full-workspace-green, or aggregate-resource gate.

## 2026-08-24 amendment: Wave67 Keynote chart-caption aggregate-budget slice (not a monolith-exit gate)

Commit `688ddbb0bd6ca4e7addc15ac8fbd12b0032fc5a9` adds one private
aggregate resource budget around the already admitted selector-first Keynote
chart-caption transaction. It also adds output-free archive sizing, prepared
ZIP reassembly, prepared PackageMetadata token/addition execution, typed limit
mapping, exact/max-minus-one integration coverage, and a ratchet that requires
the focused commit and apply paths to retain those charges. The semantic
public API and the Wave65 host cutover are unchanged; no raw-ID compatibility
surface was restored.

This is a bounded resource and publication refinement of the canonical
caption lifecycle. Unsupported future graphs, ambiguous owners, malformed
wire, hostile Metadata, and unproven dependencies remain rejected. Existing-
caption replacement uses two charged private candidates and therefore does
not establish a single-allocation package transaction. The native record is
bounded Keynote 14.4 open/render/save/close/reopen evidence with an exact
pre-native inverse, not byte parity, post-native focused Rust re-ingress,
arbitrary graph acceptance, package-wide performance, or a general
publication gate.

The `litchi-iwa` crate, remaining Keynote/Pages/Numbers compatibility hosts,
dependency edges, ordered debts, generated schemas, Buffa/Prost owners, and
monolith remain. No manifest edge, crate, remaining host, generated-schema
owner, or normal Prost owner was removed. The live checker retains the three
user-owned untracked Pages table-lock findings and no Keynote caption finding.
Wave67 therefore satisfies no crate-exit, host-exit, dependency-removal,
generated-schema-retirement, monolith-deletion, package-wide-performance, or
full-workspace-green gate.

## 2026-08-24 amendment: Wave68 Keynote chart-title host-retirement slice (not a monolith-exit gate)

Commit `e62b6fdb1` removes the raw-ID Keynote chart-title method, its
identifier-to-position wrappers, and the associated host fallback path. The
19 migrated examples and host tests now call the selector-first semantic
chart-title owner; the focused `litchi-keynote` Package behavior is unchanged.
This is a narrow host-retirement step toward the monolith boundary, not a
claim that the host is replaceable as a whole.

The `litchi-iwa` owner edge, remaining Keynote graph/chart/text/media/soundtrack
and build operations, ordered debts including debt014, compatibility-host
tests, generated schemas, Buffa/Prost owners, and the monolith remain. Wave68
does not remove a manifest edge or crate, retire a generated-schema or normal
Prost owner, or establish a host/edge/debt/monolith exit gate. No public
raw-ID API is added and no claim of full-workspace verification or package-wide
performance is made.

The bounded native evidence is the disposable Keynote 14.4 record from
`/private/tmp/litchi-wave68-chart-title-native.PGGXxv/rust-chart-title.key`
(12,796 bytes, SHA-256
`96a9070e5416fe74fe5e229a86f43733bcb0095c515fd25d5b70dd965d70d42d`) and its
same-path saved/closed/reopened normalization (156,131 bytes, SHA-256
`87893cbd99f402049d8e1eac884f8cfbd290de04dc4ac91083f097fd4afc0ac4`). Keynote
opened without repair and displayed the recorded Quarterly Results slide and
Quarterly revenue chart/caption/data. This is bounded application acceptance,
not Rust/native parity, performance/RSS, publication, or any monolith-exit
evidence.

## 2026-08-24 amendment: Wave69 shared chart-caption codec slice (not a monolith-exit gate)

Commit `cd382f95dad52fecdfdb17a7fbb1565c489e6de1` centralizes one
strict TSCH/TSD/TSP chart-caption reference edge and removes the Pages and
Numbers caption-local generated `IWorkChartArchive` decodes and lossy nested
patch helpers. The neutral hidden codec is shared with the already focused
Keynote spelling; no generated type or raw object identifier is added to a
public API.

This is a bounded shared-codec and local-decode-retirement slice. Pages and
Numbers remain compatibility-host owners for chart selection, creation and
removal graphs, themes, caption storage/placement, UUID registration,
PackageMetadata, save tokens, compression, archive publication, and native
application compatibility. Wave69 does not create a focused Pages or Numbers
chart-caption package owner and does not retire their semantic host methods.

The `litchi-iwa` crate, remaining Keynote/Pages/Numbers compatibility hosts,
manifest edges, ordered debts, generated schemas, Buffa/Prost owners, and
monolith remain. No dependency edge, crate, host, debt, generated-schema
owner, or normal Prost owner was removed. The Pages 14.4 record is bounded
open/render/save/close/reopen evidence only, not Rust/native parity,
performance/RSS, publication, full-workspace verification, or any
crate/host/edge/debt/monolith-exit gate.

## 2026-08-24 amendment: Wave70 Pages drawable-order bounded internal codec slice (not a monolith-exit gate)

Commit `9b44e20ac0d3696aeb28f94d923612157446a83c` moves the complete
`TP.DrawablesZOrderArchive` repeated `TSP.Reference` edge from generated
Prost decoding in the Pages compatibility host to the hidden strict
`litchi_iwa_protos::pages_drawable_order_codec`. The host still owns document
graph selection, semantic order validation, raw-preserving permutation,
transactional candidate/readback/reopen, and native compatibility. There is
no focused `litchi-pages` owner because the current document-wide operation
has no meaningful public selector; this slice adds no raw-ID facade or public
physical type.

The codec preserves complete reference records, interleaved root unknowns,
unknown balanced groups, and overlong unknown scalar framing while rejecting
noncanonical known framing/value forms, missing or zero identifiers,
duplicates, and non-permutation edits. Existing host serialization may
canonicalize the outer IWA object-length prefix. The native Pages record is
bounded open/render/save/reopen evidence only, not Rust/native byte parity,
performance/RSS, broad graph acceptance, or publication completeness.

The `litchi-iwa` migration host, Pages/Numbers/Keynote compatibility owners,
manifest edges, all 13 ordered debts, generated schemas, Buffa/Prost owners,
and monolith remain. Wave70 removes no crate, dependency edge, host
responsibility, debt item, generated-schema owner, or normal Prost owner; it
therefore makes no host-exit, dependency-removal, or monolith-deletion claim.

## 2026-08-25 amendment: Wave71 Pages artifact and footnote migration slice (not a monolith-exit gate)

Commit `580a5343a2c75a1c1b185a5cc8aff4a87e2a5c11` retires the public Pages raw
artifact accessor and one raw-ID host operation. The focused package keeps
exact retained ZIP bytes private and exposes `Package::write_to` with typed
partial-sink `WriteError`; the redundant public `from_archive_bytes` alias is
removed. Exact output remains a streaming sink operation only: it does not
flush, sync, rename, or atomically or durably publish a filesystem path.

`PagesEditor::set_body_footnote_text` is removed. Existing-root body-footnote
text replacement is owned by selector-first
`litchi_pages::Package::edit_body_footnote_text`; the compatibility host still
owns footnote reads and insert/remove graph lifecycle, native graph handling,
and the remaining Pages mutation surface. The section-content bridge uses one
pass `FocusedCandidateWriter` and preserves its existing `3S+2T` work
accounting.

This is a bounded API and operation retirement step, not deletion of
`litchi-iwa`. The `litchi-iwa -> litchi-pages` manifest edge, debt 017 and the
other 12 ordered debts (13 total), migration-host examples/tests, generated
schemas, Buffa/Prost owners, and monolith remain. Wave71 retires no crate,
dependency edge, debt item, host responsibility, generated-schema owner, or
normal Prost owner; it establishes no host-exit, dependency-removal, or
monolith-deletion gate. No full-workspace-green, package-wide-performance,
durable-publication, or native/Rust byte-parity claim follows.

## 2026-08-25 amendment: Wave72 shared chart-caption hardening slice (not a monolith-exit gate)

Commit `997e0bf55e8a1381296a4d446f83dd5217a15385` hardens the hidden shared
chart-caption codec and the existing compatibility-host retarget paths. All
internal codec traversals are metered; unknown overlong scalar values and
source groups remain retained while known selected fields, keys, lengths, and
group depth stay strict. Pages and Numbers preserve raw `ArchiveInfo` headers
and exact aggregate/`FieldInfo` references, rejecting shared or misattributed
`CaptionInfo`/storage ownership atomically. Keynote charges the residual codec
report and candidate reopen in its private `CaptionBudget`, with typed output
and allocation failures.

Wave72 creates no focused Pages or Numbers chart-caption package owner.
Graph/theme creation and the remaining chart lifecycle remain in
`litchi-iwa`; no public API, manifest edge, generated schema, Prost/Buffa
owner, host responsibility, or ordered debt is retired. The 13-debt ledger,
all format compatibility hosts, and the monolith remain, so this slice makes
no crate-exit, host-exit, dependency-removal, debt-retirement,
generated-schema-retirement, or monolith-deletion claim. The duplicate known
`MessageInfo` scalar canonicality and arbitrary unmodeled payload/reference
parity items remain P2 follow-ups rather than completed deletion gates.

The Pages 14.4 record is bounded application open/render/save/close/reopen
acceptance only; it is not Rust/native byte parity, a performance/RSS result,
a Numbers native result, or a full publication/workspace gate.

## 2026-08-25 amendment: Wave73 focused Keynote movie-caption owner (not a monolith-exit gate)

Commit `6ed7ee4e277c78a99f1b779a2b5530d7267beddf` moves existing file-movie
caption reads and `Some -> Some` replacement behind the selector-first
`litchi-keynote::Package` facade and a hidden strict movie-caption codec. The
focused owner proves the selected slide/movie/caption graph, exact Metadata
component identities and save tokens for every changed native component,
preview invalidation, full candidate reopen/locality, and reversible exact
artifacts without publishing native IDs.

This is intentionally not full movie-title/caption lifecycle ownership.
Caption creation/removal and title CRUD remain compatibility-host graph
operations; the legacy raw-ID host methods, native graph builders, metadata
UUID/watermark paths, media management, and normalized-producer compatibility
remain in `litchi-iwa`. The hidden Buffa projection, generated schemas, normal
Prost owners, manifest edges, all ordered debts, and the migration monolith
remain. No crate, dependency edge, host, debt, generated-schema owner, or
normal Prost owner is retired, so Wave73 establishes no host-exit,
dependency-removal, debt-retirement, or monolith-deletion gate.

The Keynote 14.4 record is bounded application open/render/save/close/reopen
evidence for the pre-Keynote focused candidate. Keynote's normalized artifact
does not pass the focused strict re-ingress policy; no Rust/native parity,
arbitrary graph acceptance, performance/RSS, durable publication,
full-workspace, or monolith-exit claim follows.

## 2026-08-25 amendment: Wave74 full Keynote movie-caption owner (not a monolith-exit gate)

Implementation commit `40c3d0b217e4b17304efca24c6a81b9d1240246b`
moves the canonical Keynote file-movie caption lifecycle behind the
selector-first `litchi-keynote::Package` owner. The focused owner now performs
read, no-op, replacement, stand-in-to-four-object creation, active-caption-to-
fresh-stand-in removal, Metadata UUID/watermark/save-token updates, preview and
locality verification, exact patch inverse, and candidate reopen. The
`litchi-iwa` compatibility host delegates through typed slide/movie selectors;
its raw caption ID methods and legacy caption graph fallback are retired.

This is a bounded owner slice, not a monolith-exit gate. Movie-title CRUD and
broader movie/media/build/theme compatibility remain in `litchi-iwa`.
Noncanonical, cross-component, shared, or unproven caption graphs continue to
fail closed rather than expanding the accepted producer graph. The
`litchi-iwa` to `litchi-keynote` dependency edge, its ordered migration debt,
the compatibility host, generated schema/Buffa/Prost owners, and the IWA
monolith remain; no crate, dependency edge, debt item, generated owner, or
normal Prost edge is deleted by Wave74.

The fresh Keynote 14.4 record establishes bounded creation/removal
open/render/save/close/reopen acceptance and Rust structural reread of the
normalized packages. It does not establish byte parity, arbitrary native graph
acceptance, performance/RSS, full-workspace health, or a host/edge/debt/
monolith retirement gate.

## 2026-08-25 amendment: Wave75 full Keynote movie-title owner (not a monolith-exit gate)

Implementation commit `ad7fc362036fa81c70c3e8ac373d7901e7689179` moves the
canonical file-movie title lifecycle into the selector-first Keynote package:
read, no-op, replacement, stand-in-to-four-object creation,
active-to-fresh-stand-in removal, Metadata UUID/watermark/save-token updates,
preview/locality verification, exact inverse, and candidate reopen. The
compatibility host now routes title reads and mutations through typed
slide/movie selectors; its raw-ID title mutators and legacy title graph
fallback are retired. Wave74 remains the separate full caption owner.

This is a bounded vertical owner slice, not deletion of the migration host.
Movie caption, media, build, theme, and other Keynote compatibility work must
remain at their recorded owners until independently migrated. The
`litchi-iwa -> litchi-keynote` dependency edge and debt 014 remain open because
the ledger exit requires completion of all Keynote package editing plus its
examples and tests. The generated schema, hidden Buffa projection, normal
Prost owners, remaining host responsibilities, and IWA monolith remain; no
crate, dependency edge, debt item, generated owner, or normal Prost edge is
deleted by Wave75.

The fresh Keynote 14.4 record provides bounded title replacement, removal, and
creation open/render/save/close/reopen acceptance. The normalized active
title/caption packages do not satisfy the narrow focused re-ingress policy, so
this is not Rust/native byte parity, arbitrary graph acceptance,
performance/RSS, durable publication, full-workspace health, or a host,
dependency-edge, debt, or monolith-exit gate.

## 2026-08-25 amendment: Wave76 Pages body-table dimension bounded owner slice (not a monolith-exit gate)

Implementation commit `e624ebab09a891eb9d5921f6ebf8c0e601bf5dff` moves the
selector-first body-table display-dimension edge into `litchi-pages`. The
focused owner provides semantic row-height and column-width read/edit/apply
operations through `BodyTableSelector` and
`table::dimension::{Dimension, Points, Size}`, with exact inverse, preview,
candidate-reopen, and locality checks. The hidden neutral codec owns the
strict raw-preserving header-bucket edge; the package owns rooted table/model/
bucket selection, cross-component ownership, and lock/dependency validation.
Malformed, duplicate, ambiguous, shared, resource-owned, and otherwise
unproven routes fail closed. No native identifier, archive route, wire value,
generated schema type, Buffa view, or Prost value crosses the focused facade.

The six raw-ID PagesEditor dimension methods and tracked `tables/layout.rs`
owner are retired by this slice. This is only the body-table display-size edge:
table appearance, hidden axes, footnotes, table topology/resize machinery,
media and other Pages compatibility responsibilities remain in `litchi-iwa`.
The `litchi-iwa -> litchi-pages` dependency edge, debt 017, `litchi-iwa` and
`PagesEditor`, remaining host responsibilities, generated schema/Buffa/Prost
owners, and the IWA monolith remain. No crate-exit, dependency-removal,
debt-retirement, manifest, generated-schema, normal-Prost, or monolith-
deletion gate is satisfied.

The Pages 14.4 record is bounded open/render/save/close/reopen evidence for
the two display-size transitions, with exact pre-native inverse artifacts and
strict reread of the normalized packages. Pages normalized the artifacts and
displayed values that differ in presentation from the strict semantic points
where documented; this is not Rust/native byte parity, a performance/RSS
result, durable publication evidence, or a full-workspace gate. The live
boundary checker retains the three user-owned untracked Pages table-lock
findings, which are unrelated to this owner.

## 2026-08-25 amendment: Wave77 Pages body-footnote lifecycle bounded owner slice (not a monolith-exit gate)

Implementation commit `5dc2ab5337cb61b72f83e369d818829c355d7141` moves the
bounded body-footnote lifecycle into the selector-first Pages package: ordered
read, checked insertion, selected text/custom-mark edit, complete selected
removal, Metadata UUID/external/watermark/save-token transitions, preview and
locality verification, candidate reopen, patch conflict detection, and exact
inverse artifacts. The hidden graph codec remains raw-preserving and strict;
native identities and generated or wire representations remain private.

The compatibility examples and host lifecycle path now route through that
owner. This is not a general body-text graph-reclamation authority: ordinary
body replacement still refuses cleanup when the exact aggregate,
`FieldInfo`, and Metadata ownership cannot be attributed. Section lifecycle,
header/footer and other text graphs, table topology and appearance, comments,
bookmarks, media/drawables, and the remaining Pages compatibility surface stay
at their recorded owners.

Debt 017, the `litchi-iwa -> litchi-pages` dependency edge, `litchi-iwa` and
`PagesEditor`, remaining host responsibilities, generated schema/Buffa/Prost
owners, and the IWA monolith remain. No crate, dependency edge, debt item,
production manifest dependency, generated-schema owner, normal Prost owner,
or monolith-exit gate is retired by Wave77.

The Pages 14.4 insertion/removal record provides bounded
open/render/save/close/reopen acceptance, exact pre-native inverses, and strict
Rust reread of both normalized artifacts. Pages normalized the ZIP bytes; this
is not Rust/native byte parity, arbitrary graph acceptance, a performance/RSS
result, durable publication evidence, full-workspace health, or a host,
dependency-edge, debt, or monolith-exit gate.

## 2026-08-25 amendment: Wave78 Pages header/footer text bounded owner slice (not a monolith-exit gate)

Implementation commit `1596d5106ee42fe9238e8d63cc55d39c494d59c3` moves the
existing-root Pages header/footer text slice into the selector-first package.
`HeaderFooterSelector` selects a section, first/even/odd template, header or
footer role, and checked zero-based slot. `Package` owns semantic reads,
`Some -> Some` replacement, clear-to-empty, exact inverse/apply, Metadata
root/selected-token handling, preview invalidation, strict codec/text-wire
rewriting, candidate reopen, and locality verification. Native IDs, raw
storage types, archive names, generated messages, wire values, Prost, and
Buffa views remain private.

This is not a header/footer graph lifecycle owner. Slot/template creation or
removal, inheritance and page-variant settings, number attachments beyond
the focused route, section lifecycle, media, annotations, table topology and
appearance, body/other text graphs, and unresolved shared/resource ownership
remain at `litchi-iwa` or their recorded owners and fail closed when they are
not provable. The hidden `pages_header_footer_codec` and text-wire owner
preserve supported unknown raw fields and reject malformed or conflicting
known ownership.

Debt 017, the `litchi-iwa -> litchi-pages` dependency edge, `litchi-iwa` and
`PagesEditor`, remaining Pages host responsibilities, generated-schema,
Buffa, and normal Prost owners, and the IWA monolith remain. No crate,
dependency edge, debt item, manifest dependency, generated-schema owner,
normal Prost owner, host-exit, or monolith-exit gate is retired by Wave78.

The Pages 14.4 record provides bounded header/footer open/render/save/close/
reopen evidence, exact pre-native inverse artifacts, and strict normalized
Rust reread. Pages normalized the ZIP bytes; this is not Rust/native byte
parity, arbitrary producer-graph acceptance, a performance/RSS result,
durable publication evidence, or a full-workspace claim.

## 2026-08-25 amendment: Wave79 Pages section-text host-retirement bounded owner slice (not a monolith-exit gate)

Implementation commit `507193d3c2ea7c6f6939f47189be5a7b425661c0`
moves existing rooted Pages section-text reads and checked edits into the
selector-first package and retires the four raw-ID host methods plus their
fallback writer. `Package` now owns semantic selection, strict text-wire
splicing, section-boundary shifts, exact patch/inverse/apply behavior,
aggregate resource accounting, candidate reopen, topology, and locality for
this narrow storage-only slice. Native IDs, archive/member routes, wire and
generated values, Prost, and Buffa views remain private.

This is not a general Pages text or section-graph owner. Rootless/nested text,
section insertion/removal, whole-body flattening across structural markers,
header/footer lifecycle, footnote graphs outside their focused owner, tables,
annotations, drawables/media, and the remaining Pages compatibility surface
stay at their recorded owners. The known aggregate-only type-10015
drawables-z-order consumer is retained; ambiguous, shared, aliased, data,
`FieldInfo`, dependent-marker, and otherwise unattributed ownership fails
closed. Metadata/save tokens and previews remain exact because this operation
does not allocate/cull objects or invalidate layout.

Debt 017, the `litchi-iwa -> litchi-pages` dependency edge, `litchi-iwa` and
`PagesEditor`, remaining Pages host responsibilities, generated-schema,
Buffa, and normal Prost owners, and the IWA monolith remain. No crate,
dependency edge, debt item, manifest dependency, generated-schema owner,
normal Prost owner, migration-host, or monolith-exit gate is retired by
Wave79.

The Pages 14.4.1 record provides bounded set/clear open/render/save/close/
reopen evidence, exact pre-native inverses, and strict normalized Rust reread.
Pages normalized both ZIP artifacts; this is not Rust/native byte parity,
arbitrary producer-graph acceptance, a performance/RSS result, durable
publication evidence, or a full-workspace claim.

## 2026-08-25 amendment: Wave80 Numbers table-appearance bounded owner slice (not a monolith-exit gate)

Implementation commit `bf01576c090cb508ac0596a5faac01951ef502d0`
moves existing rooted Numbers table-appearance reads and copy-on-write
replacement into the selector-first package. The semantic facade contains no
native IDs or physical IWA types. The strict hidden codec and package jointly
own direct model-style selection, bounded inheritance, source-preserving
stylesheet registry append, UUID/watermark and selected-token transition,
staged-private atomic publication, candidate reopen, exact patch/inverse, and
locality for this narrow slice.

This is not general style or table graph ownership. Direct field 3 remains
authoritative while any nonzero field-48 preset is opaque and byte-preserved;
preset-only graphs, preset/network/style lifecycle, reset/cull, aliases,
cross-component stylesheet ownership, table topology/content, and unsupported
shared or malformed resources fail closed or remain at their recorded owners.
The mutating Numbers editor wrapper is retired, but its table inventory keeps
a read-only compatibility fallback, and the shared appearance host remains for
Pages and Keynote.

Debt 015, the `litchi-iwa -> litchi-numbers` dependency edge, debt 017 and the
Pages edge, `litchi-iwa`, remaining migration-host responsibilities,
generated-schema, Buffa, and normal Prost owners, and the IWA monolith remain.
No crate, dependency edge, debt item, production manifest dependency,
generated-schema owner, normal Prost owner, migration host, or monolith-exit
gate is retired by Wave80.

The Numbers 14.4 record provides bounded appearance open/render/save/close/
reopen evidence, a byte-exact pre-native inverse, and strict normalized Rust
reread. Numbers normalized the ZIP bytes; this is not Rust/native byte parity,
arbitrary producer-graph acceptance, a performance/RSS result, durable
publication evidence, or a full-workspace claim.

## 2026-08-25 amendment: Wave81 Numbers table-cell Pop-Up Menu bounded owner slice (not a monolith-exit gate)

Implementation commit `4ea040b4f9cf97f61ae9a690c6ae592b1b7ac567`
moves admitted rooted Numbers table-cell Pop-Up Menu reads and lifecycle edits
into the selector-first package. The archive-free facade contains no native ID
or physical IWA type. The strict hidden codec, storage-wire owner, metadata
owner, and package jointly own exact model/cell-spec parsing, table-list and
BNC refcount proof, same-component copy-on-write/reuse/create/reset/final-cull,
identifier and save-token transitions, staged-private atomic publication,
candidate reopen, exact patch/inverse, and locality for this narrow slice.

This is not general data-format, control-cell, table, or cross-component graph
ownership. Cross-component aliases or dependencies, segmented/malformed lists,
unproven inbound references, inconsistent refcounts, ambiguous/current-
versioned metadata, other control formats, and unsupported producer graphs
fail closed. The generic Numbers data-format host delegates the Pop-Up Menu
branch to the package, while Pages/Keynote adapters and the remaining Numbers
compatibility surface retain their recorded responsibilities.

Debt 015, the `litchi-iwa -> litchi-numbers` dependency edge, debt 017 and the
Pages edge, `litchi-iwa`, remaining migration hosts, generated-schema, Buffa,
and normal Prost owners, and the IWA monolith remain. No crate, dependency
edge, debt item, production manifest dependency, generated-schema owner,
normal Prost owner, migration host, or monolith-exit gate is retired by
Wave81.

The Numbers 14.4 source record proves a native Pop-Up Menu can be authored,
saved, closed, and reopened without repair. Strict Rust read refused its
cross-component graph with `UnsupportedDependency`, so no Rust candidate,
inverse, normalized reread, or native mutation acceptance is claimed. This is
not Rust/native byte parity, arbitrary producer-graph acceptance, a
performance/RSS result, durable publication evidence, or a full-workspace
claim.

## 2026-08-25 amendment: Wave82 Keynote movie-playback bounded owner slice (not a monolith-exit gate)

Implementation commit `666ee3ec3be5d7574ebb9324154b550fde83a5f1`
moves admitted existing file-backed Keynote slide-movie playback reads and
scalar replacement into the selector-first package. The archive-free facade
contains no native ID or physical IWA type. The strict hidden codec and package
jointly own raw-preserving playback projection, unique rooted same-component
selection, operation-local budgeting, staged-private atomic publication,
candidate validation/readback, exact patch/inverse, and locality for this
narrow slice. Metadata and previews remain byte-exact.

This is not movie graph, media/poster, geometry, title/caption, build, audio,
or cross-component ownership. Duplicate, aliased, non-file, parent/owner-
ambiguous, malformed, or unsupported producer graphs fail closed. The former
raw-ID Keynote playback methods are retired, while the shared media adapter and
remaining Keynote, Pages, and Numbers compatibility responsibilities retain
their recorded owners.

Debt 014, the `litchi-iwa -> litchi-keynote` dependency edge, debts 015 and
017, `litchi-iwa`, remaining migration hosts, generated-schema, Buffa, and
normal Prost owners, and the IWA monolith remain. No crate, dependency edge,
debt item, production manifest dependency, generated-schema owner, normal
Prost owner, migration host, or monolith-exit gate is retired by Wave82.

The Keynote 14.4 record proves bounded native open/render/control/save/close/
reopen of the Rust candidate, an exact pre-native inverse, and strict semantic
reread of the Keynote-normalized artifact. Keynote normalized the ZIP bytes;
this is not Rust/native byte parity, arbitrary producer-graph acceptance,
performance/RSS, durable publication evidence, or a full-workspace claim.

## 2026-08-25 amendment: Wave83 Numbers table-sort bounded owner slice (not a monolith-exit gate)

Implementation commit `20aca0ede7817e7fbb630338dbb4bcc852d7d500`
moves admitted rooted Numbers persisted sort-order reads and configuration
edits into the selector-first package. The archive-free facade contains no
native ID or physical IWA type. The strict hidden codec and package jointly
own raw-preserving field-44 projection, rooted selection, operation-local
budgeting, staged-private atomic publication, candidate readback/locality,
exact patch/inverse, and preview/metadata preservation for this narrow slice.

This is not physical row sorting, table storage, row/column UID, formula,
comment, border, table topology, Pages/Keynote sort, or general graph
ownership. The physical `apply_table_sort_order*` executors and remaining
compatibility responsibilities stay in `litchi-iwa`; unsupported, malformed,
aliased, locked-change, or otherwise unproven graphs fail closed.

Debt 015, the `litchi-iwa -> litchi-numbers` dependency edge, debts 014 and
017, `litchi-iwa`, remaining migration hosts, generated-schema, Buffa, and
normal Prost owners, and the IWA monolith remain. No crate, dependency edge,
debt item, production manifest dependency, generated-schema owner, normal
Prost owner, migration host, or monolith-exit gate is retired by Wave83.

The Numbers 14.4 record proves bounded native open/render/sort-control/save/
close/reopen of the Rust candidate, an exact pre-native inverse, and strict
semantic reread of the Numbers-normalized artifact. Numbers normalized the ZIP
bytes; this is not Rust/native byte parity, physical-sort acceptance,
arbitrary producer-graph acceptance, performance/RSS, durable publication
evidence, or a full-workspace claim.

## 2026-08-26 amendment: Wave84 Pages body-table sort bounded owner slice (not a monolith-exit gate)

Implementation commit `31d5081ca6cd56256e463bee0019e4c7241d6df1` moves
admitted rooted Pages persisted body-table sort configuration into the
selector-first `litchi-pages` package. The archive-free facade contains only
`BodyTableSelector` and `table::sort::{Order, Rule, Scope, ColumnIndex,
Direction, RowRange}` values. The strict shared neutral codec and package
jointly own raw-preserving field-44 projection, lock refusal, operation-local
budgeting, staged-private atomic publication, candidate readback/locality,
exact patch/inverse, and prepared reassembly. Field 45 is strictly validated
but remains opaque and byte-exact; Metadata, UUIDs, save tokens, and previews
are not mutated.

This is persisted sort-order ownership only. It is not physical row sorting,
table storage or UID movement, formula/comment/cell editing, table topology,
table creation/removal, Pages/Keynote sort ownership, or general graph repair.
The physical PagesEditor apply/reorder executor and remaining compatibility
responsibilities stay in `litchi-iwa`; unsupported, malformed, aliased, locked,
or otherwise unproven graphs fail closed. Native UI rule distinction was
withheld because the empty-cell fixture left the sort commands disabled.

Debt 017, the `litchi-iwa -> litchi-pages` dependency edge, `litchi-iwa`,
remaining migration hosts, generated-schema, Buffa, and normal Prost owners,
and the IWA monolith remain. No crate, dependency edge, debt item, production
manifest dependency, generated-schema owner, normal Prost owner, migration
host, or monolith-exit gate is retired by Wave84. The native record is bounded
open/render/save/close/reopen evidence only, with normalized ZIP bytes; it is
not Rust/native byte parity, physical-sort acceptance, performance/RSS,
durable publication, or a full-workspace claim.

## 2026-08-26 amendment: Wave85 Numbers unified cell-control bounded owner slice (not a monolith-exit gate)

Implementation commit `8f804fdc65f5d99a61d8b73503353901edf87488` moves
admitted same-component Numbers Checkbox, Star Rating, Slider, Stepper, and
compatibility Pop-Up Menu control transactions into the selector-first
package. The archive-free facade contains no native ID or physical IWA value.
The strict hidden codec and package jointly own mixed control projection,
BNC/list refcount authority, copy-on-write/cull, metadata transitions,
operation-local budgeting, staged-private atomic publication, candidate
readback/locality, and exact patch/inverse artifacts for this bounded slice.

This is not arbitrary scalar data-format, table topology, Pages/Keynote
control, or cross-component format/control ownership. The dedicated Numbers
editor method families are retired, while the generic focused bridge, private
source-built compatibility readers/resets, shared Pages/Keynote adapters, and
remaining table graph responsibilities stay in `litchi-iwa`. Malformed,
segmented, aliased, cross-component, locked, ownership-ambiguous, or otherwise
unproven graphs fail closed.

Debt 015, the `litchi-iwa -> litchi-numbers` dependency edge, debts 014 and
017, `litchi-iwa`, remaining migration hosts, generated-schema, Buffa, and
normal Prost owners, and the IWA monolith remain. No crate, dependency edge,
debt item, production manifest dependency, generated-schema owner, normal
Prost owner, migration host, or monolith-exit gate is retired by Wave85.

Numbers 14.4 opened, saved, closed, and reopened the fresh UI-authored source
without repair, but the focused owner correctly refused its cross-component
format/control graph. No Rust candidate or inverse was published, so native
mutation acceptance is withheld. This inventory is not Rust/native byte
parity, publication acceptance, performance/RSS, durable publication, or a
full-workspace claim.

## 2026-08-26 amendment: Wave86 Numbers unified cell-control split-read bounded owner slice (not a monolith-exit gate)

Implementation commit `3dfe506f4febe7f389db4f60b2988430bbd6038e` extends the
bounded Numbers unified cell-control owner only to strict reads and
byte-exact no-ops for selected split-component format/control graphs. The
selector-first facade remains
`Package::{table_cell_control_format, edit_table_cell_control_format,
apply_table_cell_control_format}` with archive-free control values.

The selected split edge must have current/effective locator and external-edge
proof; versioned, conflicting, effective-locator-disagreeing, and
physical-alias targets fail closed. A single cached `RegistryFacts`/physical
census serves the logical read. Selected `TableModel` sidecar refs require
one aggregate occurrence; explicit `FieldInfo` must be unique and
`ObjectReference` typed, while producer-omitted `FieldInfo` is accepted and
no exact path claim is made. The root Document/TableInfo-to-
CalculationEngine/TableModel metadata edge is not owned or proven, and
same-component graphs do not require external-edge inspection.

Opaque inbound refs are accepted for read/no-op. Every changed split-component
route, including popup-only routes, invokes the pre-publication refusal and
preserves source bytes. No cross-component COW, UUID/save-token transition,
candidate/locality verification, inverse, or successful split write is owned;
no zero-allocation-before-refusal claim follows. This is not table topology,
general scalar formatting, Pages/Keynote control, or arbitrary native graph
ownership.

Debt 015, the `litchi-iwa -> litchi-numbers` dependency edge, debt 017, the
`litchi-iwa -> litchi-pages` edge, `litchi-iwa`, remaining migration hosts,
generated-schema, Buffa, and normal Prost owners, and the IWA monolith remain.
No crate, dependency edge, debt item, production manifest dependency,
generated-schema owner, normal Prost owner, migration host, or monolith-exit
gate is retired by Wave86.

The current Numbers 14.4 artifact is read-only provenance: it opened without
repair and exposed the split control graph, while Rust matched four controls.
There is no native mutation, candidate, inverse, normalization, byte-parity,
performance/RSS, durable-publication, host/edge/debt-exit, or full-workspace
claim in this owner slice.

## 2026-08-26 amendment: Wave87 Keynote movie-geometry bounded owner slice (not a monolith-exit gate)

Implementation commit `e11a4cc993cf29e5524745f2fa51dd3dd7d20b3e` moves only
the admitted existing same-component file-backed movie position and displayed
size transaction into the selector-first `litchi-keynote` package. The
archive-free facade is `MovieGeometry` with
`Package::{slide_movie_geometry, edit_slide_movie_geometry,
apply_slide_movie_geometry}`, `SlideSelector`, and `MovieSelector`. The
hidden geometry codec and package jointly own strict source-preserving
projection, preview invalidation, operation-local budgets, staged-private
publication, candidate reopen, semantic readback, exact locality, and
patch/inverse artifacts.

The slice does not own movie creation/removal, media or poster assets,
captions, titles, playback, builds, metadata/UUID/save-token lifecycle,
generic drawable geometry, or cross-component repair. Native angle and flags
remain preserved compatibility state, and legacy flip/original-size-restore
compatibility remains retained at the `litchi-iwa` host/adapter boundary.
Malformed, non-file, aliased, cross-component, or otherwise unproven graphs
fail closed. The native normalized artifact's strict Rust reread returned
`InvalidSource`; it is not a native/Rust parity or publication gate.

Debt 014, the `litchi-iwa -> litchi-keynote` dependency edge, debts 015 and
017, the Pages edge, `litchi-iwa` and its remaining host responsibilities,
generated-schema, Buffa, and normal Prost owners, migration hosts, and the
IWA monolith remain. No crate, dependency edge, debt item, production
manifest dependency, generated-schema owner, normal Prost owner, host-exit,
or monolith-exit gate is retired by Wave87.

## 2026-08-26 amendment: Wave88 Numbers split cell-control write bounded owner slice (not a monolith-exit gate)

Implementation commit `d1e11c7f218583becc85be7fa426650aa0113de4`
extends only the existing unified Numbers cell-control package owner from
split-component read/no-op to successful Checkbox, Star Rating, Slider, and
Stepper write transactions. The selector-first facade and archive-free
`CellControl` values are unchanged. The bounded owner now proves exact
current/effective component and metadata authority, performs scalar-control
copy-on-write/create/reset/final-cull transitions, advances the required root
and current-component save tokens, and verifies candidate semantic readback,
preview invalidation, locality, patches, and exact inverses.

Cross-component Pop-Up Menu mutation remains fail-closed and source-exact;
the popup model lifecycle is not broadened. The slice does not own general
scalar formatting, table topology, Pages/Keynote controls, arbitrary
cross-component repair, or a public native graph. Its staged-private resource
policy is not a single global allocation-free preflight or a performance/RSS
claim. Numbers-normalized native artifacts are not Rust/native byte parity.

Debt 015, the `litchi-iwa -> litchi-numbers` dependency edge, debts 014 and
017, the Pages edge, `litchi-iwa` and its remaining host responsibilities,
generated-schema, Buffa, and normal Prost owners, migration hosts, and the IWA
monolith remain. No crate, dependency edge, debt item, production manifest
dependency, generated-schema owner, normal Prost owner, host-exit, or
monolith-exit gate is retired by Wave88.

## 2026-08-26 amendment: Wave90 Numbers split Pop-Up Menu bounded owner slice (not a monolith-exit gate)

Implementation base `e867003d80fb9cc9373f2f659e2b51c3264de06b` moves only
the admitted split-component Pop-Up Menu lifecycle from fail-closed refusal
to a successful, bounded `litchi-numbers` owner slice. The selector-first
facade remains archive-free. The private owner proves member ownership,
current/effective locators, exact external edges, metadata UUID and
save-token transitions, BNC/list references and refcounts, copy-on-write,
reuse/create/reset/final-cull behavior, prepared reassembly, candidate
readback/locality, and exact patch/inverse behavior.

Native acceptance now supplies semantic persistence evidence for the
admitted fixture: Numbers opened the 78,984-byte Rust candidate without
repair, displayed `None`, `Wave90-Low`, and `Wave90-High`, and preserved
the choices across save, close, and exact-path reopen. Strict Rust reread
accepted the 135,726-byte Numbers-normalized candidate. The 135,226-byte
source and exact Rust inverse share SHA-256
`b830e8981cd9cd0ede3f125506b6b7076e33b7cc79110a8a59696270d005afc9`;
the native-normalized artifact is not Rust/native byte parity.

This is not a monolith-exit gate. The split route remains limited to its
validated graph shape and does not own arbitrary producer graphs, broad
Numbers/table formats, table topology, Pages/Keynote controls, native
identifiers, durable publication, performance/RSS, or general graph repair.
No host method, crate, dependency edge, debt item, production manifest
dependency, migration host, generated-schema/Prost/Buffa owner, or monolith
gate is retired by Wave90. Debts 014, 015, 016, and 017 remain open.

## 2026-08-26 amendment: Wave91 Numbers comment-reply read bounded owner slice (not a monolith-exit gate)

Implementation base `e8749ca9e5320c5bf99c17967bf850f7c6e14d6c` moves only
the ID-free, selector-first projection of source-ordered direct comment replies
into the focused Numbers package boundary. The slice validates the recognized
comment graph with strict codecs and package-wide ownership checks, rejects
external/deprecated-type/nested/cyclic/aliased/malformed routes, and redacts
authored content from debug output.

This does not move reply mutation, native identity, author/UUID lifecycle,
comment-list repair, candidate publication, or the compatibility editor. It
does not claim native direct-reply acceptance, opaque future-owner coverage,
one aggregate operation-budget pass, full-workspace green status, or
performance/RSS behavior.

This is not a monolith-exit gate. Debt 015, the
`litchi-iwa -> litchi-numbers` edge, the remaining hosts and debts,
generated-schema/Prost/Buffa owners, and the IWA monolith remain. No crate,
edge, debt item, production manifest dependency, host, or monolith owner is
retired by Wave91.

## 2026-08-26 amendment: Wave92 comment-reply rewrite enabling primitive (not a monolith-exit gate)

Implementation commit `65bdac3ece028a997c73dc82c4d49ef7839ff002`
adds only a private prepared codec primitive for ordered reply-reference
append, checked replacement, and checked removal. It preserves admitted raw
source framing and returns bounded codec-local prepare and execution reports.
No selector-first reply mutation package owner, native graph lifecycle, host
retirement, or native direct-reply acceptance is added.

The future bounded owner still needs rooted sheet/table/cell selection,
CommentStorage list and BNC key/refcount census, exact aggregate/FieldInfo and
global inbound authority, shared-thread COW, author and metadata UUID/token
transitions, archive/compression/reassembly preflight, candidate reopen and
locality, patch/inverse artifacts, and native evidence. Until then the
identity-bearing `litchi-iwa` compatibility mutation surface remains.

This is not a monolith-exit gate. Debt 015, the
`litchi-iwa -> litchi-numbers` edge, the remaining hosts and debts,
generated-schema/normal Prost/Buffa owners, and the IWA monolith remain. No
crate, edge, debt item, production manifest dependency, host, or monolith
owner is retired by Wave92.

## 2026-08-26 amendment: Wave93 Numbers comment-reply bounded owner slice (not a monolith-exit gate)

Implementation commit `21004a78ec4c6c7d8436424de270ad6ed6051eb0`
moves only the admitted existing-root, same-member, direct-leaf reply lifecycle
into the focused Numbers package. The archive-free facade owns semantic
ordinals and authored text; the private owner owns strict rooted graph,
refcount, author, ArchiveInfo, Metadata, COW/cull, resource, publication,
candidate verification, locality, patch, and inverse behavior.

The slice deliberately excludes root-comment creation, shared reply leaves,
nested/cyclic replies, segmented lists, cross-member writes, author creation,
opaque graph repair, Pages/Keynote comments, and arbitrary producer graphs.
The `litchi-iwa` Numbers methods retain a deprecated compatibility fallback
where native identity or an unsupported graph prevents focused delegation.
The staged-private resource policy is not a single global allocation-free
preflight or a performance/RSS claim, and the inert Numbers Reply UI means
native direct-reply acceptance remains withheld.

This is not a monolith-exit gate. Debt 015, the
`litchi-iwa -> litchi-numbers` edge, the remaining hosts and debts,
generated-schema/normal Prost/Buffa owners, and the IWA monolith remain. The
authoritative inventory stays at 64 packages, 239 internal dependency
declarations, and 13 ordered migration debts. No crate, edge, debt item,
production manifest dependency, compatibility host, generated-schema owner,
normal Prost/Buffa owner, or monolith gate is retired by Wave93.

## 2026-08-26 amendment: Numbers raw-ID comment-clear retirement (not a monolith-exit gate)

Commit `4f7f2c16c` deletes one deprecated raw-ID mutation and its
fallback bridge from `litchi-iwa`: table-cell comment clear is now exercised
through `litchi_numbers::Package`, semantic sheet/table selectors, and checked
cell positions. The low-level cleanup helper is restricted to host regression
tests, and the migrated example no longer imports `NumbersEditor` or accepts a
native table identifier.

This is a real host-surface reduction but not a crate or edge exit. The raw-ID
comment read/replacement compatibility surface, reply compatibility, other
Numbers graph operations, Pages/Keynote hosts, debt 015, the
`litchi-iwa -> litchi-numbers` edge, remaining debts, normal generated/Prost/
Buffa owners, and the IWA monolith remain. Native, resource/performance,
package-count, dependency-count, debt-exit, and monolith-exit claims are
withheld beyond the recorded scoped gates.

## 2026-08-26 amendment: strict existing-graph Numbers root-comment creation (not a monolith-exit gate)

The focused Numbers `Package::set_table_cell_comment` transaction now admits
`None -> Some` when the selected cell already has a physical BNC slot and the
rooted table already owns one strict, unsegmented comment list plus one
resolvable author. The private owner allocates the storage identifier and UUID
against physical and current/versioned Metadata namespaces, emits a canonical
Buffa `CommentStorageArchive` leaf, appends the exact list entry and
`ArchiveInfo [3,key]` edge, rewrites only the selected tile cell, advances the
owning current-component and root save tokens, removes previews, reopens the
candidate, and retains exact apply/inverse artifacts. The strict integration
suite also covers sibling preservation, conflict, missing author, unknown
Metadata, and cross-component fail-closed behavior.

The slice does not create a missing comment list, author, tile, or sparse BNC
slot; it does not admit segmented or cross-component creation, generated
message publication, native Numbers UI acceptance, or a single global
allocation-free resource preflight. Those graphs continue to fail closed.
This is therefore a semantic capability increment, not a crate, dependency
edge, debt, compatibility-host, normal Prost/Buffa owner, or monolith exit.

## 2026-08-26 amendment: Numbers empty author-registry root creation (not a monolith-exit gate)

The focused root-comment transaction now also admits a rooted same-member
table whose comment list and physical target cell already exist and whose
canonical annotation-author storage is present but empty. The owner allocates
one author and one comment identifier together, emits the local type-212
author plus canonical Buffa comment leaf, appends the author-storage and
comment-list references with exact ArchiveInfo paths, registers both UUIDs,
advances one current-component token and the root watermark/save token, and
retains exact candidate reopen and inverse restoration.

This does not create a missing type-213 author storage, missing comment list,
tile, or sparse BNC slot, and it does not admit segmented or cross-component
creation. The deprecated raw-ID host remains necessary for those legacy
blank/generated-table topologies. Native Numbers UI acceptance, a single
global allocation-free resource preflight, performance/RSS, crate/dependency
edge/debt retirement, normal Prost/Buffa ownership, and monolith exit remain
withheld.

## 2026-08-26 amendment: Numbers missing comment-list creation (not a monolith-exit gate)

The focused root-comment owner now also admits a rooted same-member table with
an existing physical target cell and canonical empty author registry but no
comment-list object or DataStore field-19 reference. A prepared lazy codec
inserts the nested `comment_storage_table` reference without re-encoding the
TableModel/DataStore envelope. The owner allocates list, comment, and author
identifiers together; creates the type-6005 list with exact `[3,key]`
ArchiveInfo, appends the model `[4,19]` authority edge, registers all three
UUIDs, advances one component token and the root watermark/save token, reopens
the candidate, and retains exact inverse restoration.

The slice still requires a canonical type-213 author storage, existing tile,
row, and BNC cell slot. Sparse-cell/tile creation, missing author-storage
creation, segmented/cross-component graphs, arbitrary repair, and native UI
acceptance remain unsupported. The raw-ID compatibility host, debt 015,
remaining dependency edges/debts, performance/RSS and single-global-preflight
claims, normal Prost/Buffa owners, and monolith exit remain.
