# 0362ul87b — shared Web Extensions and Ribbon publication boundary

Status: implementation complete; source-only until the coordinator grants a
scoped validation slot. This lane closes the semantic `&mut OpcPackage`
surfaces identified by 0362ul87. Byte/XML codecs and the physical OPC graph
substrate remain unchanged. Review repair 87ba made staging borrowed/lazy and
made Ribbon replay prove the complete retained dependency closure.

The shared package API has one source-preserving route for each family:

* `web::Edit::new(&OpcPackage)` stages `put` and `remove` operations and
  `commit` returns `web::Commit = litchi_core::edit::Commit<OpcPackage,
  web::Patch>`; `web::Patch::apply(&OpcPackage)` returns the same triad and
  never mutates its receiver.
* `ribbon::Edit::new(&OpcPackage)` stages `put` and `remove` operations and
  `commit` returns one shared `ribbon::Commit` for either or both fixed family
  slots. There are no public one-operation Ribbon publishers. The staged edit
  clones a package only on its first real mutation; the private in-place graph
  writers remain an implementation detail.
* `web::apply_durable` takes `&OpcPackage` and returns the ordinary Web
  Extensions commit. Durable replay therefore keeps the bounded deferred
  planner and does not publish through a raw mutable package parameter.

Owner callers replace their staged candidate with `commit.into_package()`;
the source package remains unchanged. Exact no-ops return the source package
with `Diagnostics::UNCHANGED`, preserving signatures and retained lexical
captures. Edit options are intersected with the package's retained
`SourceLimits`. Changed signed packages continue to fail before publication.
Web edits retain their requested Strict/Transitional profile, including empty
relationship graphs. Ribbon patches compare family identity, raw root and
part `.rels`, content types, payload allocations and the complete opaque
dependency closure before using OPC preservation APIs for replay. Inverse and
independent-copy/conflict tests cover those guards.

The implementation migrates DOCX, PPTX, XLSX and XLSB owner routes and the
shared/common tests in the same change. It does not add aliases or deprecated
forwarders. Cargo, rustdoc, native qualification and performance validation
remain coordinator work after the complete source tree is frozen.
