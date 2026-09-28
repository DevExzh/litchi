# 0806 source review

**Subsequent correction:** `quality-1/` disproved the original visibility
assessment below. The candidate narrowed the OLE helper's `BytesStartExt` and
`CheckedAttributes` declarations from `pub` to `pub(crate)`, breaking existing
crypto imports. `visibility-amendment-review.md` records the independent
five-helper audit and exact restoration. The statements below retain the
initial review for traceability; they do not override that correction or the
required production gates.

Result: the exact 0805 candidate archive and the 0806 public probe pass source
review. This review does not certify a Cargo build, a before qualification, a
native or allocation capture, a profile run, or production adoption. Those
gates remain root-owned and must consume the frozen source and report oracles.

## Candidate identity

The six `candidate/before` files are byte-identical to the current production
files at `e3ff267ee3454e71d66f177f54f3cd05e0d9cce5`. The six `candidate/after`
files are the exact after archive from the qualified 0805 preflight. The
candidate manifest records the 0805 derived manifest, seal, and preflight
decision; no new production implementation is introduced by this packet.

The five helper copies are the same implementation after the existing mirror
normalization removes the OPC visibility spelling and copied test-module path.
The shared OPC test source is the only test-source change. The source set is
limited to the five helper paths and that shared test path. The current
production-to-archive SHA checks are:

| Production path | Before archive SHA-256 |
| --- | --- |
| `crates/litchi-ole-common/src/xml_attributes.rs` | `067bdb3d1b7b8023742bef74f84cd121c2bcfc33cba4dbb61f47f7c953536d64` |
| `crates/litchi-opc/src/xml_attributes.rs` | `6c7e5aeb66f1ca420f5e11af0bfa307ae1b5fd6aa29b4d3b5af57d6b9ec6fb56` |
| `crates/litchi-opc/src/xml_attributes/tests.rs` | `6c7d024783556aab85af7c142953fe0a09a848e9f6afde814f0f83523b29f098` |
| `crates/litchi-sign/src/xml_attributes.rs` | `7e389b5a474c1856673a98fb7954cebd84685d691390dbfd7d3dbafde28a42de` |
| `crates/litchi-xldm/src/xml_attributes.rs` | `7e389b5a474c1856673a98fb7954cebd84685d691390dbfd7d3dbafde28a42de` |
| `crates/xml-minifier/src/xml_attributes.rs` | `7e389b5a474c1856673a98fb7954cebd84685d691390dbfd7d3dbafde28a42de` |

## Attribute iterator semantics

The after source keeps the public iterator and error contract. The unchecked
extension method still exposes quick-xml's unchecked iterator, while
`checked_attributes` returns quick-xml-compatible items through the first
error and is fused after an error or end. The candidate starts quick-xml's
parser with duplicate checking disabled, then uses raw key-prefix preflight
before asking the parser to consume a value. This preserves duplicate-before-
malformed-value precedence and the original byte offsets without asking the
parser to revisit a yielded value.

The private phases cover the first item, the second-item boundary, the
two-item short path, the bounded borrowed-name linear path, the ordered path,
and `Done`. The linear path stores at most quick-xml's 32-name boundary in a
fixed array. A unique later name is returned once; only a successful unique
33rd name with a remaining tail seeds the ordered map from borrowed keys.
There is no reconstructed iterator and no replay of earlier values. Duplicate
lookups retain the first occurrence's position. The test archive covers the
short/linear/ordered boundaries, malformed tails, duplicate precedence,
leading-equals keys, long values, clone continuation, fusion, and post-error
exhaustion.

The exact-empty decision is made on the first request. This changes private
derived `Debug` state from the production construction state to the candidate's
`First` state before the first `next` call; the private phase and nested parser
state therefore have different debug text at construction and transitions.
That compatibility caveat is retained explicitly. No public method, trait,
field, or visibility is added or removed by the candidate.

The candidate no longer calls the retained `unchecked_attributes` trait method
from production iterator code; the method is still part of the existing trait
and is exercised by the test module. The four private mirror copies may report
an unused-method warning when compiled outside `cfg(test)` under the workspace
warning policy. This is an execution gate for the root quality run, not a
speculative source amendment. The exact after archive remains immutable until
that gate supplies its result.

The inherited source comments still describe the implementation as an 0802
candidate and contain an outdated immediate-map sentence. The actual source,
archive hashes, and tests describe the bounded linear transition reviewed
above; the comment mismatch is retained as provenance and does not authorize
editing the frozen candidate.

## Public probe and fixture contract

The 0806 probe is a source-only extension of the qualified public PPTX
harness. It preserves the 0794 `capture_region_0793` profile owner, the
`capture`, `commit`, and `lifecycle` timing scopes, operation-scoped allocator
region, opaque result use, and post-clock semantic verification. Package
capture and setup remain outside `commit`; the lifecycle clock includes the
documented capture, edit, commit, publication, and serialization calls. No
allocator result is used as a latency claim, and no profile result is used as
a latency or RSS claim.

The original fifteen rows keep the same deterministic source-text generator,
3x4, 12x8, and 100x100 dimensions, six unknown-URI vendor entries, marker,
rewrite paths, semantic text, and readback identities. The vendor value is
constructed through a new internal field but produces the same legacy
`litchi-perf-0785-{note}` attribute bytes. Their before qualification must
match the sealed 0792 source and output identities independently of any new
row.

The added `valid-4attr` row uses the vendor dimensions, 12 slides by 8 text
tags. Every `a:t` start tag receives exactly these four distinct, quoted,
namespaced attributes:

| Prefix and URI | Local name | Value |
| --- | --- | --- |
| `lx1`, `urn:litchi:perf:0806:extension:one` | `probeOne` | `litchi-perf-0806-valid-4attr-one` |
| `lx2`, `urn:litchi:perf:0806:extension:two` | `probeTwo` | `litchi-perf-0806-valid-4attr-two` |
| `lx3`, `urn:litchi:perf:0806:extension:three` | `probeThree` | `litchi-perf-0806-valid-4attr-three` |
| `lx4`, `urn:litchi:perf:0806:extension:four` | `probeFour` | `litchi-perf-0806-valid-4attr-four` |

The four `xmlns:lxN` declarations are inserted on the slide's `<p:sld>` root
start tag. The source validator checks distinct prefixes, URIs, names, and
values, valid XML-name/value characters, and double-quoted assignments. The
post-save oracle checks every slide's root start tag and rejects a local or
duplicate declaration, requires exactly four global `xmlns:lx` declarations,
requires each exact declaration once on the root and once in the slide blob,
parses every `a:t` start tag, and requires exactly the four expected
name/value pairs with no duplicates. It also checks the complete semantic
text and records text-tag, attribute, value, declaration, URI, name, and value
counts in the sample verification object.

The packet analyzer must accept a report only when it matches the 0806 probe
schema/tool and, for `valid-4attr`, verifies those metadata and sample fields:
12 slides, 8 text tags per slide, 4 declarations per slide, 4 attributes per
text tag, 96 total text tags, 384 attribute/value occurrences, and the exact
four URI/name/value lists. It must require the new three-row before-only
qualification oracle and keep the fifteen-row 0792 parity check separate. A
source review cannot substitute for that strict report replay.

## Disposition

The exact candidate is eligible for the frozen 0806 workflow gates from source
review. The candidate remains preflight-only and production adoption remains
false until the before qualification, candidate quality gate, all eighteen
fresh public rows, resource guards, profile custody, cross-format veto, raw
audits, cleanup, and final seal pass. No historical timing is pooled with the
0806 measurements.
