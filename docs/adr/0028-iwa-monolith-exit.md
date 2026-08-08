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
