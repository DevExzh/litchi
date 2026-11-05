# 0362ul4 — exact inventory semantic categories

This tooling-only leaf records three narrow ways an inventory row can be
semantic rather than a format-package API debt row.  A semantic category is
not a waiver: it does not repair a source item, defer it, or approve a final
disposition.  The all-68 source capture is still pending, so this leaf adds no
final disposition rows.

## Exact classifications

`tools/validate_api_disposition.py` accepts these decisions only for the exact
inventory identities reviewed in [0362ul2](0362ul2-iwa-protos-raw-dispositions.md):

* `AST_VALUE`: `litchi-formula` / `mut_owner` /
  `litchi_formula::Formula::{set_root,set_display_style}`.
* `WIRE_REQUEST`: six `litchi-iwa-protos` / `edit_type` codec request paths
  from the Keynote media-lifecycle and Numbers cell-storage codecs.
* `WIRE_SCHEMA`: the 13 `litchi-iwa-protos::tswp` protobuf paths in 0362ul2,
  each in its observed `edit_type` and `transaction_type` inventory rows.

The validator compares the complete `(crate, category, path)` identity.  It
does not accept a crate-wide or module-prefix exception, a removed alias, or
an unknown sibling.  Evidence and owner remain mandatory, one-to-one rows
remain mandatory, and `DEFERRED_REDACTION_22` remains the only deferral.

Focused tests cover every reviewed path plus wrong crates, categories,
unknown siblings, nested paths and removed aliases.  The build-free check was:

```text
python3 -m unittest -q tools.test_validate_api_disposition
```

No Cargo command, source capture, API-debt closure, or final acceptance claim
is made here.
