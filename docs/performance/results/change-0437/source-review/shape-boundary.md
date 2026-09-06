# Fresh ODP matched shape boundary

The initial 64/8,192/32,768-slide proposal exceeded an existing buffered
publication default. `PackageWriter::validate_authored_xml` invokes
`xml_minifier::audit::Limits::default()`, whose aggregate attribute ceiling is
250,000. This is below the immutable 1,000,000-attribute ceiling. The format
Builder does not expose a caller override for that authored-member audit.

The retained `checks/buffered-harness-shape-boundary.json` release test receipt
passes six focused tests, including construction of 32,768 deterministic titled
slides. The actual refusal is:

> XML publication rejected for 'content.xml': XML Attributes limit 250000 exceeded by 250001 at byte 8846661

The matched performance corpus therefore uses tiny=64, medium=4,096, and
large=8,192. The 32,768 case is retained as a refusal test, not a completed
creation measurement. Production limits remain unchanged. Streaming capacity
under separately selected explicit limits would require a separately scoped
measurement; it cannot replace a missing matched buffered result.

This correction happened before any timed pilot, preparatory profile, formal
measurement, or production optimization. The initial fixture identities are
retained as `fixture-identities-initial-shapes.json` to explain the abandoned
shape proposal; they are deterministic text identities, not accepted archive
or performance results.
