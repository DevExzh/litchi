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
rejects observable in-place mutation. Protobuf failures retain their source
chain behind a format-owned, content-free wrapper.

This is a cutover prerequisite, not the cutover. Reopening a legacy
`litchi_iwa::Document` through `Package` would still break in-memory and
directory-backed inputs, violate immutable snapshot semantics, duplicate
physical parsing, and impose strict rooted failures on the historical global
API. A shared immutable catalog/source coordinator and package-wide
compatibility budgets for cells, sidecars, text, formulas, and error mapping
remain required before deleting `structured/numbers.rs`. The remaining table
model, formula-owner, sidecar, and AST Prost decoders still need focused
pre-decode envelopes or projections.
