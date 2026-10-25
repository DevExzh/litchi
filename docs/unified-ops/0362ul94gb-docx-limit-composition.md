# 0362ul94gb — preserve DOCX document-property limit identity

This follow-up supersedes the mapping claim in `0362ul94ga`. A document
properties limit is a distinct shared family budget, so the DOCX adapter now
converts it exactly to
`litchi_ooxml_common::LimitKind::Properties(properties::LimitKind)`.
The nested value carries the property family identity without guessing a core
or XML codec dimension and remains lossless if the properties family grows.

The common and properties limit kinds expose their own resource descriptions,
and the common `Properties` arm delegates to the nested kind. The `litchi-ops`
regression exercises all six currently known property kinds and checks the
nested kind, observed amount, maximum, display description, downstream
`limit_exceeded` code, and retained source chain for each conversion.

This correction is source-only. No Cargo, rustdoc, runtime, or performance
command was run; the composed tree must pass the coordinator's serialized
consumer gate.
