# 0362ul87b — shared Web Extensions and Ribbon publication boundary

Status: implementation complete; source-only until the coordinator grants a
scoped validation slot. This lane closes the semantic `&mut OpcPackage`
surfaces identified by 0362ul87. Byte/XML codecs and the physical OPC graph
substrate remain unchanged.

The shared package API has one source-preserving route for each family:

* `web::Edit::new(&OpcPackage)` stages `put` and `remove` operations and
  `commit` returns `web::Commit = litchi_core::edit::Commit<OpcPackage,
  web::Patch>`; `web::Patch::apply(&OpcPackage)` returns the same triad and
  never mutates its receiver.
* `ribbon::Edit::new(&OpcPackage)` stages `put` and `remove` operations and
  `commit` returns one shared `ribbon::Commit` for either or both fixed family
  slots. The immutable `ribbon::put` and `ribbon::remove` helpers are the
  one-operation convenience route; their source-checked
  `ribbon::Patch::apply` also returns a commit. The in-place graph writers are
  private implementation helpers.
* `web::apply_durable` takes `&OpcPackage` and returns the ordinary Web
  Extensions commit. Durable replay therefore keeps the bounded deferred
  planner and does not publish through a raw mutable package parameter.

Owner callers replace their staged candidate with `commit.into_package()`;
the source package remains unchanged. Exact no-ops return the source package
with `Diagnostics::UNCHANGED`, preserving signatures and retained lexical
captures. Changed signed packages continue to fail before publication. Patch
replay validates the captured family graph and applies only the selected
semantic closure, retaining unrelated package parts and payload allocations.

The implementation migrates DOCX, PPTX, XLSX and XLSB owner routes and the
shared/common tests in the same change. It does not add aliases or deprecated
forwarders. Cargo, rustdoc, native qualification and performance validation
remain coordinator work after the complete source tree is frozen.
