# 0362n — Remaining ODF raw XML policies

- Reviewed trunk: `e98e5878cb`; Part A candidate: `fd5f513df6`.
- Status: required engineering work; not a conformance exception.

The flat-document repair in 0362hd/0362hda does not close the shared ODF raw
XML surface. Source review found the following public ingress points still
selecting implicit bounds or accepting no caller policy:

| Surface in `crates/litchi-odf-common/src/` | Required follow-up |
| --- | --- |
| `core/resolved_reader.rs`: `ResolvedReader::from_xml` | Require a raw XML policy. Preserve normalized namespace resolution, zero-copy events and fallible reservation; retain caller bounds through event reads. |
| `core/metadata/codec.rs`: `Metadata::from_xml` | Require limits, preserve typed actual failures, and migrate metadata readers and edits in all ODF consumers. |
| `package/codec.rs`: `parse_manifest`, and `core/manifest.rs`: `Manifest::parse` | Thread explicit bounds through the shared observer scan and both public projections; retain encryption and namespace correctness. |
| `calculation/codec.rs`: `parse` and `parse_with_report` | Replace implicit XML/event limits with required policy and typed observations; preserve the same tolerance report. |
| `core/xml/{part,content,styles,meta}.rs`: byte/vector constructors | Bound raw storage admission explicitly and preserve the allocation-free owned-buffer path; distinguish styles and content policies. |
| `chart/reader.rs`, `style/{data,master}` and `drawing/resources` public XML parsers | Complete the same raw-policy audit and migrate every affected caller in each source commit. |

One concrete error-fidelity gap is visible in `ResolvedReader::push`: reaching
its fixed namespace scope depth returns `NamespaceError::TooManyDeclarations`
with the depth ceiling. Active namespace bytes and bindings also have fixed
internal ceilings. The migration must distinguish these dimensions and report
their actual observations, without inferring kinds from error strings.

Existing immutable package constructors accepting `OpenOptions` are not raw
API violations merely because their input is bytes. Conversely, a parser is
not exempt simply because its name lacks `_with_limits`: the current debt
inventory detects that spelling but cannot prove policy propagation. Scalar
datatype conversions and authored output models need a separate surface
classification rather than a blanket parser rename.

Every repair must preserve exact XML and tolerance behavior, check strict
profiles at retained-reader boundaries, migrate all consumers, register
appropriate glossary coverage, and include matched allocation/native evidence
for changed hot paths. The final whole-workspace inventory and gate remain
required; this review does not mark any row complete.
