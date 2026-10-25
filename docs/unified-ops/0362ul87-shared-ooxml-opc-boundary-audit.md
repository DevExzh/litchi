# 0362ul87 — shared OOXML semantic `OpcPackage` boundary audit

Status: source-only audit.  This record does not change Rust code, run Cargo,
or claim the lane is accepted.  It records the boundary that a follow-up wave
must migrate atomically.

## Basis and identity

The inspected source is the shared-OOXML tree at integration commit
`4a62cdfb035cd43cbd0a28c8d1d67fbf03e19e41`.  That commit has no source diff
from its source parent; the source identities below were checked with
`git hash-object` in the assembled worktree:

| source | blob |
| --- | --- |
| `crates/litchi-ooxml-common/src/custom_xml/package.rs` | `2b7ddc1819796c98cdcd1bd9e1dc4977dc9aa15b` |
| `crates/litchi-ooxml-common/src/custom/package.rs` | `906389da82be82fa601436ef436f9c1477553e0a` |
| `crates/litchi-ooxml-common/src/properties.rs` | `aebce492959bffce4077ecd3c6e12df652be4ec7` |
| `crates/litchi-ooxml-common/src/properties/app/package.rs` | `c12a5143416f4417426ec908ddea81b6b2dda296` |
| `crates/litchi-ooxml-common/src/properties/write.rs` | `80cf46fc9d12c549ad18c066a79140764dc6d2c1` |
| `crates/litchi-ooxml-common/src/properties/transaction.rs` | `afacd43cca8702509b66cc2b7dc62b1c2140c0f5` |
| `crates/litchi-ooxml-common/src/ribbon/package.rs` | `29c5a8b01ecc1ba75bf1ab69d9a1a37591c6345c` |
| `crates/litchi-ooxml-common/src/web/package/planning.rs` | `f20c1be71f687c7da070ebe485b66260bc34bcae` |
| `crates/litchi-ooxml-common/src/web/package/transaction.rs` | `01fb6a080d4ca266d0183aa229b8fea6f720633c` |
| `crates/litchi-ooxml-common/src/web/package/durable.rs` | `7a5e2cf09b102b712ae0f7f98f0d7765cfa6ada8` |
| `crates/litchi-ooxml-common/src/vba.rs` | `7416ae26f865f0bea9e899996328f1636aa6a8ab` |

The governing rules are 0351 §2.1 (immutable owners and the scoped
`edit_opc` escape), §2.5 (raw/codec limits), and §2.6 (migrate every
consumer), together with 0362rd §W4 boundary and §OPC substrate scope.  The
rules permit a raw physical graph and public byte/XML codecs; they do not make
a public format-semantic helper that takes `&mut OpcPackage` an exception.

## Finding: a real semantic publication boundary remains

The following caller-visible functions or methods take `&mut OpcPackage` and
perform semantic graph publication.  They are not byte-only codecs and are
outside the permitted raw `Package::edit_opc` route.  Their package mutation
is therefore a concrete 0362ul87 follow-up, not a justified exemption.

| public surface (current path and line) | graph family | current exposure |
| --- | --- | --- |
| `custom_xml::add` (`custom_xml/package.rs:81`, re-exported by `custom_xml/mod.rs`) | custom XML part, relationship and optional properties part | public free mutator |
| `custom::Props::write`, `write_for` (`custom/package.rs:115,169`) | custom-properties part and relationship | public methods on a public model |
| `properties::write` (`properties.rs:411`), `properties::clear` (`properties/write.rs:57`) | core-properties part | public free mutators |
| `properties::Slot::flush`, `stage` (`properties.rs:535,552`) | staged core-properties publication | public methods on a public slot (`Slot` is doc-hidden, not private) |
| `properties::app::write`, `clear` (`properties/app/package.rs:200,251`) | extended-properties part | public free mutators from a public module |
| `properties::Patch::apply` (`properties/transaction.rs:1968`) | metadata transaction graph | public patch application against the raw package |
| `ribbon::put`, `remove` (`ribbon/package.rs:77,162`, re-exported by `ribbon/mod.rs`) | Ribbon part, relationships and images | public free mutators |
| `web::put`, `remove` (`web/package/planning.rs:30,624`, re-exported through `web/package/*`) | persisted web-extension/task-pane graph | public free mutators |
| `web::Patch::apply` (`web/package/transaction.rs:202`) | persisted web-extension/task-pane graph | public patch application against the raw package |
| `web::apply_durable` (`web/package/durable.rs:169`) | durable web-extension/task-pane graph | public free mutator |
| `vba::store_project_graph`, `remove_project_graph` (`vba.rs:129,237`, feature-gated module) | VBA part and relationship graph | public free mutators |

`custom_xml::discover`, `properties`/`app` readers, Ribbon `load`, web
readers/planners, and VBA readers are projections or plans and are not
included in this finding merely because they mention `OpcPackage`.

## Consumer census for the migration

The current source has owner routes that already clone or stage a candidate,
which is the correct place to preserve atomic publication while the shared
helper surface is narrowed.  The following production and externally visible
consumers must be migrated in the same implementation commit; common tests
and examples are included rather than treated as exceptions.

* Custom XML: DOCX `package/package/data_stores.rs:240` calls
  `custom_xml::add` on its candidate; the DOCX `edit_custom_xml` family is the
  owner route.  Shared custom XML tests exercise the direct function.
* Core and extended properties: DOCX
  `package/package/metadata.rs:145-150` uses core `write`/`clear` and the
  metadata patch on a candidate.  The tracked example
  `crates/litchi/examples/core_props_office.rs:62-93` and
  `tools/perf-baseline/src/lib.rs:62109` call core `properties::write`
  directly.  Common core-property tests and the application-property tests
  call both public families.  `properties::Slot` and the transaction patch
  are also exercised by the shared transaction tests.
* Ribbon: DOCX `package/model.rs:860,884` and XLSB `ribbon.rs:106,125`
  call the shared `put`/`remove` from their owner edits.  Common Ribbon
  integration tests call the public functions directly.
* Web/task panes: XLSX `task_panes/edit.rs:125-134,179` calls shared
  `put`/`remove` on the staged candidate; DOCX
  `package/package/task_panes.rs:71-74`, PPTX `web/edit.rs:70,87`, and
  XLSB `task_panes.rs:175` apply the shared patch.  XLSX workbook tests and
  common web-extension tests call the public functions.  The shared durable
  helper is publicly re-exported and must be dispositioned even though no
  direct production callsite was found in this source grep; the existing DOCX
  owner durable tests exercise the corresponding owner route.
* VBA: DOCX `vba_project.rs:438,467`, PPTX
  `presentation/embedded/vba/package.rs:76,96`, and XLSX
  `active_x/vba.rs:130,149` call the shared store/remove graph helpers;
  each has an owner edit route.  The shared VBA tests also call them.
* Custom properties: the direct shared tests and
  `crates/litchi-xlsx/tests/package_signature_policy.rs:1190` exercise the
  public `Props::write_for`; the owner metadata routes must remain the
  supported publication path after migration.

This is a boundary migration, not permission to remove behavior.  The wave
must retain signed-package refusal ordering, atomicity, exact XML/part
fidelity, configured limits, and the existing owner-route performance.  It
must also audit facade/ops/CLI/Python, examples, benchmarks, fuzz targets and
all affected tests before closing §2.6; no public alias or deprecated
forwarder is acceptable.

## What is not this finding

The physical `litchi_opc::OpcPackage` graph methods remain substrate methods;
their names alone do not make them format-owner APIs.  Public byte/XML
helpers such as the Custom XML validation/read/write codec exports remain
codec surfaces when they take values/bytes and explicit `SourceLimits`, and
are not changed by this record.  Conversely, calling a graph mutation a
“staging helper” does not make its public `&mut OpcPackage` signature safe:
the current functions above still publish directly and must be routed through
the owner edit/commit shape (or made private while all consumers migrate).

No Cargo, rustdoc, native qualification, performance or zero-debt claim was
made for this source-only audit.
