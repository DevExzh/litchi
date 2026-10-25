# 0362ul87b — shared Web Extensions and Ribbon publication boundary

Status: implementation complete; source-only until the coordinator grants a
scoped validation slot. This lane closes the semantic `&mut OpcPackage`
surfaces identified by 0362ul87. Byte/XML codecs and the physical OPC graph
substrate remain unchanged. Review repairs 87ba and 87bb made staging
borrowed/lazy, made Ribbon replay prove and restore the complete retained
dependency closure (including binary assets), and made empty Web task-pane
root profile declarations part of conformance matching.

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

## F5 compile receipt and source repair

The first serialized F5 retry was intentionally stopped at the first failed
row and its evidence is preserved outside the checkout. It ran the exact
source commit `e06bc961fc88110d229a5312be95946f9b452015` with the 33-row
plan and the approved isolated runner. The formatter row passed. The next
`check-common` row failed before any later row could run:

* receipt: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul87b-f5-retry-e06bc961fc/result.json`
  (SHA-256 `1b5c2fb2a8c763d9bd4d1070ed0c1e418fd42913e6d4a5d21e360c3622a83e85`);
* failed compiler log: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul87b-f5-retry-e06bc961fc/01-check-common-attempt-actual-f5-01.log`
  (SHA-256 `8c9c059f8bbbfdce72f9e4cd139362918c05aae46cb22918d18ab2642020ee04`);
* row 1 exited `101` after 38.72 seconds, with a 158,643,465-byte target;
  the runner recorded no live-target or ownership violation.

The diagnostics were confined to the shared common crate: Ribbon replay
passed a content-types byte proof to the relationship-token comparator,
four diagnostics paths were unnecessarily qualified, and both Ribbon and
Web `const` limit intersections called the non-const `Ord::min` methods on
the pinned toolchain. This repair keeps relationship token tuples, compares
the content-types retained bytes directly beside its semantic token, uses
the existing module diagnostics re-exports, and implements the intersections
with const-safe `min_usize` and saturating `as_usize` helpers. The next F5
freeze must use the new repair commit and a fresh empty target; this failed
receipt remains bound to the pre-repair source and is not overwritten.

The repairs were committed as `c429e4fdd7024042a4a3d79e5095764d8801a0c0`.
The next serialized F5 run reached the common-crate integration tests and
found one stale consumer assertion before any later row ran:

* receipt: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul87b-f5-c429e4fdd7/result.json`
  (SHA-256 `4fd668a9e8afe2793b07fa11f222b8ceb5077a4c9111f0eae7b7cf17cb6b89b5`);
* failed compiler log: `/home/zhuhe/code/litchi-worktrees/.coord/0362ul87b-f5-c429e4fdd7/01-check-common-attempt-actual-f5-01.log`
  (SHA-256 `f597bc14dbeb308e52dbef1a9e486c8f16a1b321de392a2f6b6586d377e6df91`);
* row 1 exited `101` after 35.27 seconds with a 162,579,023-byte target;
  the exact error was the integration test's private `Panes::panes` field
  access at `web_extension_packages.rs:84`.

This follow-up changes only that consumer to the existing public
`Panes::len()` accessor. No public field or compatibility alias is added;
the remaining `.panes` accesses are crate-internal implementation or unit
tests within the owning common crate.
