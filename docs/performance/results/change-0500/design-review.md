# 0500 design review before implementation

The existing `Edit::replace_body_paragraph_texts` is the correct format-owned
public seam. `source_backed::Package::edit_document` already returns this edit,
and `Operation::ReplaceParagraphText` already supports main-document source
publication. No provider or archive implementation type needs a new public path.

The managed scalar implementation resolves current projected text, validates
immutable base text against its operation ledger, prepares a final ledger, and
reconstructs all final operations through original `SourceXmlPart` proofs. The
batch should share the reconstruction suffix: one splice publication, one
finish, one managed candidate snapshot, and readback for all final operations.
Calling the complete scalar setter in a loop would retain the measured repeated
work and would not implement the intended optimization.

Inputs remain nonempty, unique, strictly increasing, and bounded. Existing
operations must be retained, updated, or removed according to current and base
text; reverting the last operation restores the exact base owner. An all-current
batch does not rebuild. A failed final selector, text validation, admission,
source proof, parse, or readback cannot replace any part of the original edit.

Admit final operation metadata and strings before cloning or growth. Plan,
fragment, scan, and candidate overlap remain charged while the old edit stays
live. Memory, Objects, and Depth release through owners and RAII; Work and
InputBytes remain consumed. Preserve cancellation and source-version fences.
The managed source-backed publication proof must survive: converting through
an owned rewritten XML helper is insufficient.

Before production changes, the existing managed transaction integration suite
passed all 36 tests. It covers arbitrary-order scalar edits, updates/reverts,
source identity, signed no-ops, inverse artifacts, resource and cancellation
refusals, and partial sink failures. New tests must exercise the batch itself;
this design review is not implementation acceptance or performance evidence.
