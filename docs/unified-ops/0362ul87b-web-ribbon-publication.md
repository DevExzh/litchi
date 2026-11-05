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

* receipt: `(deleted-coord-artifact)/result.json`
  (SHA-256 `1b5c2fb2a8c763d9bd4d1070ed0c1e418fd42913e6d4a5d21e360c3622a83e85`);
* failed compiler log: `(deleted-coord-artifact)/01-check-common-attempt-actual-f5-01.log`
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

* receipt: `(deleted-coord-artifact)/result.json`
  (SHA-256 `4fd668a9e8afe2793b07fa11f222b8ceb5077a4c9111f0eae7b7cf17cb6b89b5`);
* failed compiler log: `(deleted-coord-artifact)/01-check-common-attempt-actual-f5-01.log`
  (SHA-256 `f597bc14dbeb308e52dbef1a9e486c8f16a1b321de392a2f6b6586d377e6df91`);
* row 1 exited `101` after 35.27 seconds with a 162,579,023-byte target;
  the exact error was the integration test's private `Panes::panes` field
  access at `web_extension_packages.rs:84`.

This follow-up changes only that consumer to the existing public
`Panes::len()` accessor. No public field or compatibility alias is added;
the remaining `.panes` accesses are crate-internal implementation or unit
tests within the owning common crate.

The fresh retry at source commit `85a4b3e618adb75611ee8f6004b66cf660bf2414`
was granted the serialized validation slot and stopped at the first compiler
failure. Its prior zero-row isolation refusal remains untouched:

* retry receipt: `(deleted-coord-artifact)/result.json`
  (SHA-256 `7802d2e8951664dc91bbf5271a3346b328cb15a09b961499f2fe5d31a92180dc`);
* rows: `(deleted-coord-artifact)/rows.json`
  (SHA-256 `2e8d7ceb06a2c199c30959fe0b7b86f2c31e03934d81a82697e0c561b824d5a0`);
* failed compiler log: `(deleted-coord-artifact)/02-check-leaves-attempt-actual-f5-02.log`
  (SHA-256 `c4661133fae0e50573eb1a84832ba67555aa129a128c3733b96d6d2ab2d47d83`);
* `fmt-scoped` and `check-common` passed; row 2 `check-leaves` exited `101`
  after 55.17 seconds with a 285,736,553-byte target. No later row ran.

The diagnostic was an XLSX task-pane adapter return-shape mismatch at
`crates/litchi-xlsx/src/task_panes/edit.rs:123`: the public `replace` verb
returns `Result<()>`, while its private staging closure correctly computes the
meaningful changed flag. The adapter now maps the successful internal result
to unit at its public boundary. DOCX task-pane `put`/`remove` and durable
routes, PPTX task-pane `set`/`remove`, and XLSB task-pane `replace`/`clear`
were audited; their public boolean results and all consumers intentionally
retain the changed signal and required no migration.

The repaired source `fe767f2285ea2c5a328ac76ef45af4dd8673ee12` was then
validated in the sole serialized slot and stopped at the first remaining
compiler diagnostic. The fe767 capture is preserved outside the checkout:

* receipt: `(deleted-coord-artifact)/result.json`
  (SHA-256 `4480421547f3afa44fa251fc662d4b8534a7322d6c2d212576a56950423ec036`);
* rows: `(deleted-coord-artifact)/rows.json`
  (SHA-256 `d4048e4290868f3bc8a35e3c933f6ca7241c1e5d03cf23e2d209f8bad92ffeea`);
* failed compiler log: `(deleted-coord-artifact)/02-check-leaves-attempt-actual-f5-02.log`
  (SHA-256 `1abe0a9524a9e77b60fd72c632f802091700aa381235bb9c94ccff6ee3b1cfd0`);
* `fmt-scoped` and `check-common` passed; row 2 `check-leaves` exited `101`
  after 63.64 seconds with a 306,737,160-byte target. No later row ran.

The remaining diagnostics were three DOCX typed-owner assignments: the
shared `Package::opc` field is `SharedOpc`, while the shared Web/Ribbon
commit publishes an `OpcPackage`. The assignments in
`package/model.rs` (Ribbon put/remove) and `package/package/task_panes.rs`
(empty-patch replay) now use the existing `From<OpcPackage> for SharedOpc`
conversion, retaining Arc-backed package sharing and all package policies.
All other Web/Ribbon Commit consumers in the four leaves were audited:
common candidates and PPTX/XLSX/XLSB staged OPC values already assign the
published `OpcPackage`, while DOCX tests and typed wrappers assign the
published DOCX `Package`; no additional shape mismatch was found.

The beac39a4c5 source freeze passed the formatter, common check, all four
format-leaf checks, and the two facade checks. It stopped at the first Clippy
failure in the shared common crate:

* receipt: `(deleted-coord-artifact)/result.json`
  (SHA-256 `8319cf286c414d1b54cb86315f3ecadb934af39cb5c97697e2d7e6272c0946d2`);
* rows: `(deleted-coord-artifact)/rows.json`
  (SHA-256 `8d00faafde7f03d7a743d280e11df8f6540690088d0d2eae17dbb3e0b88ad14e`);
* first failed log: `(deleted-coord-artifact)/05-clippy-common-attempt-actual-f5-05.log`
  (SHA-256 `d3af669c2338084ea3f52caf8417cbdc0e8d9b5e7453a3529ca4b33b1f836c5c`);
* rows 0 through 4 passed; row 5 `clippy-common` exited `101` after
  16.41 seconds with a 617,685,040-byte target. No later row ran and no
  builder remained after termination.

Clippy identified only the manual `Default` implementations for Ribbon and
Web `EditOptions`. The repair derives `Default` on both option structs and
removes those implementations; their `Limits` fields already implement
`Default` as the respective standard limits. The failed receipt remains
bound to beac39a4c5 and is not overwritten; the next F5 freeze must use this
repair commit and a fresh empty target.

## F5 glossary diagnosis and source repair

The c877dee369 retry passed rows 0 through 8 and stopped at the first
glossary row. Its evidence is preserved outside the checkout:

* receipt: `(deleted-coord-artifact)/result.json`
  (SHA-256 `8c1edf088c78367e3629c8e593730c2e73de5abd17db2c4de30a11babe2cf664`);
* rows: `(deleted-coord-artifact)/rows.json`
  (SHA-256 `00d3ff25df5f2bacae720d4a468498dacf65d41231879915fbc8a4387d2257ec`);
* failed glossary log: `(deleted-coord-artifact)/09-glossary-ooxml-common-minimal-attempt-actual-f5-09.log`
  (SHA-256 `fbf5a0137db3ffa210a560d040780c744ba45e0e64b6e8ff1ae011999ad896d9`);
* row 9 exited `1` after 16.169 seconds with a 929,942,531-byte target;
  no later row ran. The frozen rustdoc input was
  `(deleted-coord-artifact)/litchi_ooxml_common.json`
  (SHA-256 `13b3c16a395fcda07d67d0c132a7b760337fefdba2ca174640c563c85e1bd4ff`).

The six findings had two checker causes and one real API migration:

* `custom_xml::{read_props,validate_payload,write_props}` already use the
  public root re-export `litchi_opc::SourceLimits`. Dependency rustdoc exposes
  its nominal item only as `litchi_opc::limits::SourceLimits`; the checker now
  uses an explicit audited root-re-export binding to that defining path,
  remaining fail-closed for wrong middle paths and duplicate identities. No
  raw limit signature or public spelling was changed.
* Ribbon and Web `Patch::apply` already return the correct core triad
  `Result<Commit<OpcPackage, Patch>>`. The shared owner checker had compared
  that result against the `Patch` receiver; it now checks the package source
  argument for `Patch::apply` and retains exact nominal result matching.
* Web `Patch::is_empty` was a real glossary violation. It is now the required
  `is_noop`, with every common, DOCX, XLSX and XLSB consumer migrated in the
  same change and no alias left behind. Signed no-op and source-proof
  semantics are unchanged.

The checker changes have fixture coverage for the valid root re-export,
wrong-middle-path and duplicate-identity rejection, and `Patch::apply`
package-owner publication.
