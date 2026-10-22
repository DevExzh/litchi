# 0362nb — Return the raw ODF reader error boundary for repair

Source review of `747cfd072d` finds that the required XML policy now reaches
the ODT/ODS owner paths, but the public raw reader still loses its precise
limit kind at its own boundary. This candidate is not accepted as completion
of 0362n. The follow-up repair is 0362nc.

At the reviewed commit, `litchi-odf-common/src/core/resolved_reader.rs`
imports `litchi_core::{Error, Resource, ResourceLimit, Result}`. Its public
`from_xml`, `read_resolved_event` and `read_resolved_event_into` methods
therefore return the core error. Its private `RawXmlDimension` becomes a
coarse `ResourceLimit` with a scope string. `error.rs` then reconstructs
`LimitKind::RawXml` through `RawXmlLimitKind::from_scope` when a caller
explicitly converts that core error to the ODF crate error.

The new tests in `tests/resolved_reader.rs` check namespace scope, depth and
declaration kinds only after `OdfError::from(error)`. They do not demonstrate
the typed public raw contract. These findings are from the committed source;
no failing runtime or performance measurement is claimed.

The repair must return the ODF crate error directly from the raw reader,
construct `LimitKind::RawXml` from its actual enum dimension and test that
returned error before conversion. Existing legacy consumers may explicitly
convert at their own boundary. Preserve the exact observation/maximum,
namespace behavior, retained caller policy and historical owner admission.
Every affected consumer must migrate in the same change. Rerun common raw
tests, the ODT/ODS suites and affected consumer checks with retained logs.
Performance qualification remains open.
