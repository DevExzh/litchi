# 0362ul94k — OOXML leaf compile repair

This source-only follow-up starts from the committed 94j tip
`491137abe2f7b4961beaac445e463be3c7f40418` (tree
`9f4fd05f54b8b2a381e5675c47e902fde2b2bfbc`) after the common and OPC rows
passed and the leaf rows found three caller errors.  It makes no acceptance
claim and does not rerun the gate.

The DOCX data-store reference audit now iterates the `Relationships` views
through their explicit `.iter()` method for both package-level and part-level
relationships.  The XLSX signature-policy integration test consumes the
public `Package::into_plain_opc` boundary instead of calling the crate-private
`Package::opc` accessor; the test still signs the edited clear OPC graph and
exercises the same signed/no-op/removal policy.

Only rustfmt, diff, and repository size/binary guards are run for this repair;
Cargo, rustdoc, compiler, native, and performance execution remain pending.
