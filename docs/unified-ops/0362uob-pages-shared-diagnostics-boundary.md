# 0362uob — Pages shared diagnostics boundary

The crate-boundary CLI exposed a stale assumption in the Pages transaction
family audit: `metadata_edit` deliberately re-exports the shared
`litchi_core::edit::Diagnostics` and aliases the shared generic
`litchi_core::edit::Commit<Package, Patch>`, while the audit required every
canonical name to be a nominal definition in the Pages owner file.

The audit now records this as an explicit shared-type contract. For the
metadata family it requires the exact, unaliased
`pub use litchi_core::edit::Diagnostics;` statement, rejects a local duplicate
or an alias, and validates the shared core definition and its public
`Diagnostics::changed` accessor. The existing checks still require the
canonical owner and public-module exports, the generic commit alias, and all
other family names; this is not a family exemption or an alternate API.

The fixture builder and regression tests cover the accepted shared re-export
and reject both a direct alias and a multiline grouped alias, including when
the exact canonical re-export is still present. The production Rust source is
unchanged.

Checks:

```text
python3 -m unittest \
  tools.test_check_crate_boundaries.BoundaryPolicyTests.test_pages_transaction_families_accept_w1_layout_and_dispatch \
  tools.test_check_crate_boundaries.BoundaryPolicyTests.test_pages_transaction_families_reject_retired_names_aliases_and_globs
```

Both tests passed. The focused audit of the checked-in owner reports
`metadata_edit boundary clean`; the full boundary CLI is rerun in this packet.
