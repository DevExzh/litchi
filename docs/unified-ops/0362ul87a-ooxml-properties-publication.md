# 0362ul87a — OOXML properties and Custom XML semantic publication closure

Status: source-only implementation on `uo2/ooxml-properties-publication-0362ul87a`,
based on integration `66364d4fb0dc98ab04072f248cd6008f5974346f`.  This record
does not claim a Cargo build, rustdoc run, native qualification, performance
measurement, or final gate acceptance.

## Finite API inventory

The following semantic publication surfaces from 0362ul87 are closed in this
commit.  The public replacement is the shared Edit/Commit shape; no alias or
deprecated forwarder remains.

| old surface | public shape after this commit | owner/codec detail |
| --- | --- | --- |
| `custom_xml::add` | `custom_xml::Edit::insert`, then `Edit::commit` | `add` is crate-private; DOCX and shared tests publish the returned commit |
| `custom::Props::write`, `write_for` | `properties::Edit::set_custom_props`, then `Edit::commit` | the typed model and host validation remain unchanged |
| `properties::write` | `properties::Edit::set_metadata`, then `Edit::commit` | the core writer remains a private codec helper |
| `properties::clear` | `properties::Edit::clear_metadata`, then `Edit::commit` | removal still removes the selected graph, preserving no-op/signature ordering |
| `properties::Slot::flush`, `stage` | private slot implementation paths | `Slot` retains load/read/staging state capabilities; callers use the shared Edit publication boundary |
| `properties::app::write`, `clear` | `properties::Edit::set_app_props`, then `Edit::commit` | app codec/read projections remain public; graph mutators are crate-private |
| `properties::Patch::apply(&mut OpcPackage)` | `properties::Patch::apply(&OpcPackage) -> properties::Commit` | replay checks the immutable source and publishes the commit triad; there is no public `apply_to` twin |

The commit triad remains `litchi_core::edit::Commit` with package, patch and
diagnostics accessors.  `Patch::apply` checks the recorded source graph before
copy-on-write publication, preserves signed no-ops, refuses changed signed
packages before mutation, and leaves the source package unchanged.  The
test-only in-place adapter is not part of the release API and only forwards
through the immutable replay route.

## Consumer census and migrations

Every consumer identified by the boundary audit was migrated in this
implementation commit:

* DOCX `package/package/data_stores.rs` uses `custom_xml::Edit::insert` for
  Custom XML publication.
* DOCX `package/package/metadata.rs` uses `properties::Edit` for core writes
  and clears, and `families::Patch::apply` for application/custom replay.
* The common Custom XML, custom-properties, application-properties and
  transaction tests now exercise the Edit/Commit routes.  Private codec tests
  continue to test byte/XML codecs and their graph preconditions directly.
* XLSX `package_signature_policy.rs` stages custom properties through the XLSX
  metadata owner before constructing its signed fixture.
* `crates/litchi/examples/core_props_office.rs` publishes PPTX and XLSX core
  properties through the shared edit with the corresponding host policy.
* `tools/perf-baseline` stages its core metadata through the same edit route;
  its benchmark workload and measurement logic are unchanged.

No shared Ribbon, web/task-pane or VBA source was changed; those are the
separate 0362ul87b/87c lanes.  Physical OPC graph operations and public
byte/XML codecs remain available at their existing explicit-limit boundaries.

## Source-only validation

The following checks passed without invoking Cargo, rustdoc, native tools or
performance harnesses:

```text
git diff --check
rustfmt --edition 2024 --check <changed Rust source files>
```

The scoped source scan reports no public `&mut OpcPackage` mutator for the
families above and finds the sole public properties replay signature as:

```text
properties::Patch::apply(&self, package: &OpcPackage) -> Result<Commit>
custom_xml::Patch::apply(&self, package: &OpcPackage) -> Result<Commit>
```

The parent coordinator should run the scoped compile/test target and then the
full gate after the W4/W4b integration.  That later validation must cover
signed no-op and refusal ordering, exact XML/relationship fidelity, patch
conflict replay, owner host policy, retained-byte admission and the existing
DOCX/XLSX benchmark routes.
