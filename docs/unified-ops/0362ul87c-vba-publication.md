# 0362ul87c — shared OOXML VBA publication boundary

Status: source-only implementation on `uo2/ooxml-vba-publication-0362ul87c`.
This slice follows the boundary audit in 0362ul87. It has not run Cargo,
rustdoc, native qualification, or performance measurements; those require a
coordinator validation slot after the four OOXML consumers are composed.

The shared VBA graph now exposes `vba::Edit::new(&OpcPackage, Host)`,
`Edit::{store,remove,reset}`, `Edit::commit()`, and the shared `Patch` replay
family. The edit captures the selected host graph from an immutable source,
shares unchanged payloads, and publishes an `litchi_core::edit::Commit`
containing the new `OpcPackage`, exact before/after graph patch, and
content-free diagnostics. `Patch::apply(&OpcPackage)` authenticates the
recorded graph base, publishes a candidate atomically, and leaves the
recipient unchanged on refusal; `Patch::inverse()` swaps the exact graph
states. The old public free mutators are gone. Signed-package refusal, atomic
graph validation, content-type transitions, relationship exclusivity, Word's
supplemental data, and byte-preserving payload sharing remain in private
candidate helpers behind the edit family.

The DOCX, PPTX, and XLSX owner edit families continue to publish through their
existing `Edit::commit` triads. Their private staging adapters compose the
shared VBA edit and assign its committed candidate only after the operation
succeeds; the owner source remains unchanged on every refusal. The existing project
discovery and explicit `litchi_vba::Limits` payload parser are read-only and
were not widened into a raw mutation route. XLSB has its own host-local VBA
graph implementation and is unchanged by this shared OOXML slice.

Consumers migrated in this commit:

* `litchi-docx::vba_project::{store_vba_project,remove_vba_project}`;
* `litchi-pptx::presentation::embedded::vba::{store,remove}`;
* `litchi-xlsx::active_x::vba::{store,remove}`; and
* the shared graph tests, including signed refusal, dangling-target atomicity,
  and late content-type rollback.

No public semantic helper in this slice takes `&mut OpcPackage`. The physical
OPC graph methods and explicit-limit VBA reader remain codec/substrate APIs;
no alias or deprecated forwarding entry point was added.
