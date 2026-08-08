# ADR 0028: Ordered exit of the legacy IWA migration host

- Status: Accepted
- Date: 2026-08-08
- Amends: ADR 0002, ADR 0010, and the IWA record currently stored as
  `0023-iwa-index-foundation.md`

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

The duplicated ADR number on `0023-iwa-index-foundation.md` remains historical
file identity. This record amends its migration-host wording and makes concrete
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
