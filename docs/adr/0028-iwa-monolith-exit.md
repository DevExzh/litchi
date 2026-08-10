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
sanitizer campaign was executed. Consequently the host structured adapter,
its dependency, and all 17 recorded monolith debt edges remain. This amendment
does not authorize monolith deletion, claim complete Buffa laziness, or infer
edit/resave fidelity from the new source routes.

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
All 17 ordered host dependency debts remain. Slide nodes, slides, builds,
shapes, notes, tables, charts, media, other mutation paths, and portions of
semantic graph projection still use the migration host and/or generated Prost
values. Protobuf groups remain transactionally fail-closed at shared package
preflight. The unavailable sanitizer campaign, missing aggregate transaction
peak-memory option, durable JSON patch envelope, atomic filesystem save, and
remaining examples/tests/fuzz targets keep the monolith deletion gate open.

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
