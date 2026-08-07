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
