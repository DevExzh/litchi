# 0359s — legacy raw replay classification

The legacy raw owners now expose the same replay vocabulary at their public
boundaries.  A patch base mismatch keeps the owning crate's typed error:

| Owner | Base mismatch |
| --- | --- |
| OLE property sets and OLE1 objects | `litchi_cfb::OleError::InvalidFormat` |
| DataSpaces | `litchi_crypto::spaces::Error::Invalid` |
| DOC `RgDofr` and `Selsf` records | `litchi_doc::package::Error::Corrupted` |
| DOC whole-package protection patch | `litchi_doc::package::Error::InvalidFormat` |

These are malformed-input errors at the unified facade.  The shared object
patch error remains `PatchConflict` and classifies as `invalid_input`; a
positional `SourceChanged` remains `io_error`.  The public DOC record APIs use
`PatchConflict` for their semantic patch errors, and the subdocument table
patch uses the same name in place of `SourceConflict`.  Private package and
body-edit conflict names were left unchanged.

The raw replay tests cover an independently reopened source with equal
content and a different source with the typed error above.  The facade tests
also verify the malformed, patch-conflict, and source-changed classifications.
