# 0362ul94cc — Custom XML typed source-limit return

This source-only return follows the immutable 94cb implementation and the
e707 side-table correction.  It is a separate review record for the 94cc
typed-limit and active-offset fixes; it makes no acceptance or qualification
claim.

The Custom XML codec now reports source-profile depth, XML-event, and
XML-attribute-byte exhaustion through `Error::Limit` with the owning
`custom_xml::LimitKind` and the actual observed and configured maximum values.
OPC limit errors are no longer wrapped as generic `Error::Opc` values.  XML
attribute admission charges the qualified key plus raw value before decoding,
then checks the decoded value as well.  Payload validation uses the same
attribute policy.

Properties input containing MCE markup receives an early bounded physical
scan before preprocessing.  MCE output and depth ceilings are mapped back to
the Custom XML limit shape.  Active source-offset provenance is tracked in a
single fallibly reserved selected-offset table rather than by inserting one
marker comment per candidate; the ID-only proof discards processed output
after selection.  Structural rewrites consume that same table, and paired
`schemaRef` frames carry the schema-reference index rather than a byte offset,
so opaque-removal refusal does not panic.

Regression coverage asserts observed and maximum values for depth, events,
and attribute bytes, exercises payload attribute admission, and runs an MCE
rewrite under an exact source-part ceiling.  Rustfmt, diff, and repository
size guards are the only validation performed for this return.  Cargo,
rustdoc, compiler, native, and performance execution remain pending.

The existing typed-properties no-op contract is unchanged.  Its early
`before == after` branch still admits the retained byte length, while the
`Edit::set_properties` caller immediately routes the returned bytes through
`set_properties_xml`, whose `read_props` admission performs the full selected
source event, depth, and attribute checks before publishing.  No redundant
full scan was added to the no-op branch.
