# 0362ul104aa — Custom XML result return correction

The complete runtime packet on joint source `3c291e3b62` stops while compiling
common OOXML, before any test runs. The 104a alias-validation repair returns
an `Error` directly from a function returning `Result<()>`. Root's source
review missed this type mismatch. Both source/data and properties/owner
alias branches require `Err(unsupported_target(...))`.

This correction adds those two Result wrappers. The shared refusal kind,
fixed content-free reasons, validation order and pre-publication behavior
remain unchanged. No additional consumer migration is required.

Root verifies the failed terminal capture before changing source. External
`(deleted-coord-artifact)/rows.json` is 4,335 bytes,
SHA-256 `e817823231d86472dc4eb9f20ffe5b476bad522aceacf8b257dbf52d6db75869`;
`result.json` is 5,074 bytes, SHA-256
`5a3371425bc53d582a520fe7f8d6da34eeda69d872fe1e461fc32c47c2346ac8`.
The inner compiler log is 163,337 bytes, SHA-256
`5f1c4ce00030e6508381aedc105333a870736426689a80f550d21c29d54845e1`.
Target ends at 357,965,491 bytes. No runtime pass is claimed; fresh source-bound
qualification is required. Source formatting and diff checks pass.
