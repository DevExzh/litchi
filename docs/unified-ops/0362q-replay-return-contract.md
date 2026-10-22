# 0362q — Verify edit and replay return contracts

The glossary previously checked family type names and accessor names without
verifying what an edit or replay actually returns. A method returning
`Result<Package>` could therefore remain beside a registered `Commit` type.
This change verifies return types and makes unregistered owner replay visible
in the debt inventory. It changes tooling and policy; production Rust is
unchanged from `bda325e958`.

`family_methods` now checks that each local `Edit::commit` returns its
family's `Result<Commit>`. The optional `result_methods` rules name the exact
public method, Commit and generic type arguments. Registered owner replay
must publish the receiver's owner type. A snapshot Commit cannot satisfy
that rule merely because it is spelled `Commit`.

The resolver follows local aliases, generic substitutions and defaults,
re-exports and `Self`, retaining nominal type identity. It recognizes the
standard Result by identity and refuses opaque, unsupported and cyclic type
shapes. Lifetimes are erased because this is an API-shape check, not a Rust
type checker. Dependency Commit methods remain checked at their defining
crate; an unrelated foreign type renamed Commit does not qualify. Durable
wire patches may differ from the returned family's Patch; their conversion
and replay semantics remain Rust test obligations.

The DOCX minimal/full profiles register 39/40 existing conforming replay
methods. The inventory reports every remaining owner `apply`/`apply_*`
without a verified rule, even through aliases, and reports invalid edit
return shapes. Captures may now be renamed: inventory crate identity comes
from the rustdoc root instead of the JSON filename.

The [receipt](results/0362q/validation.json) retains source and artifact
hashes, exact commands and all four retained validation attempts. Final
validation passes 38 glossary/inventory tests within 1,332 tooling tests,
plus both DOCX profiles against freshly built feature-specific rustdoc.
The failed second inventory attempt exposed the filename assumption; its
log and successful reruns are retained. No performance claim is made.

The [full DOCX inventory](results/0362q/docx-full-inventory.tsv) still has
35 rows: four replay returns, eight edit types, eleven Commit types, seven
snapshot accessors, two limits twins and three edit-return findings.
Metadata, durable Ink and primary source-backed replay repairs are separate
candidates. Raw/model/projection classification and every remaining required
migration still need review; this record grants no exception. Other crates'
policy registrations, a fresh workspace inventory, glossary completion and
the final workspace gate remain open.
