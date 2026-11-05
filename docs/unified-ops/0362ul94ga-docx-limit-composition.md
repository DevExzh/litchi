# 0362ul94ga — DOCX metadata-limit composition repair

The review of `0362ul94g` found that converting a shared properties
`Error::Limit` to DOCX `Error::Other(String)` discarded the typed resource,
observed amount, and maximum. That caused the unified operations classifier to
report `internal_error` instead of `limit_exceeded`.

The DOCX adapter now converts the shared properties limit to the shared
`litchi_ooxml_common::Error::Limit`, retaining all three measurements and a
typed common `LimitKind`. The known property resources map to their matching
common budget names (`PropertiesXmlBytes`, `XmlEventCount`,
`CustomProperties`, `DocumentPropertyPartNames`, and
`SerializedPropertiesXmlBytes`); the core text budget maps to
`CorePropertyTextBytes`, and future property resources use the generic typed
`PropertiesXmlOutput` bucket. The existing DOCX common `#[source]` wrapper is
therefore retained, and the downstream `litchi-ops` regression checks both the
typed details and `limit_exceeded` classification.

No Cargo, rustdoc, runtime, or performance command was run. This follow-up is
source-only and must be rerun through the coordinator's serialized consumer
gate after composition.
